#!/usr/bin/env python3
"""
Comparador de preco SEFAZ (NFC-e) -- adaptadores Amazonas e Paraiba.

Le uma planilha de produtos com precos atuais, consulta o portal oficial de
precos por NFC-e do estado e devolve, por item, o menor preco encontrado, onde
ele esta e a que distancia. Nao emite veredito de "vale a pena": reporta
economia e distancia lado a lado.

Portais suportados
------------------
AM  Busca Preco AM        https://buscapreco.sefaz.am.gov.br
PB  Preco da Hora Paraiba https://precodahora.tcepb.tc.br   (atras de Cloudflare)
BA  Preco da Hora Bahia   https://precodahora.ba.gov.br     (mesmo codigo-base do PB,
                                                             sem anti-bot -- use para
                                                             validar o fluxo)

Uso
---
    python comparador_preco_sefaz.py --uf AM --municipio Manaus --planilha lista.xlsx
    python comparador_preco_sefaz.py --selftest        # testa a logica pura, sem rede
    python comparador_preco_sefaz.py --smoke --uf AM   # 1 consulta real, para validar acesso

Dependencias: requests, beautifulsoup4, lxml, openpyxl
    pip install requests beautifulsoup4 lxml openpyxl
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import re
import sys
import time
import unicodedata
from dataclasses import dataclass, field, asdict
from typing import Any, Iterable

UA = "Mozilla/5.0 (compatible; comparador-preco-sefaz/1.0)"
CACHE_GEO = os.path.expanduser("~/.cache/comparador_preco_sefaz_geo.json")

# --------------------------------------------------------------------------
# 1. Normalizacao de descricao e medida
# --------------------------------------------------------------------------

# Palavras de embalagem/logistica que atrapalham a busca no portal.
RUIDO = {
    "pct", "pacote", "cx", "caixa", "fd", "fardo", "und", "unid", "unidade", "un",
    "dz", "duzia", "fr", "frasco", "gf", "garrafa", "lt", "lata", "sc", "saco",
    "ref", "refil", "c", "com", "de", "da", "do", "e", "tipo", "novo",
}

# unidade -> (base, fator para a base)
UNIDADES = {
    "l": ("ml", 1000.0), "lt": ("ml", 1000.0), "litro": ("ml", 1000.0),
    "litros": ("ml", 1000.0), "ml": ("ml", 1.0),
    "kg": ("g", 1000.0), "quilo": ("g", 1000.0), "k": ("g", 1000.0),
    "g": ("g", 1.0), "gr": ("g", 1.0), "grama": ("g", 1.0), "gramas": ("g", 1.0),
    "mg": ("g", 0.001),
}

_RE_MEDIDA = re.compile(
    r"(?<![\w,\.])(\d{1,5}(?:[.,]\d{1,3})?)\s*"
    r"(ml|l|lt|litros?|kg|k|g|gr|gramas?|grama|mg|quilo)(?![a-z])",
    re.IGNORECASE,
)
# "12x500ml", "6 x 1L" -> multipack
_RE_MULTI = re.compile(r"(\d{1,3})\s*[x\*]\s*(\d{1,5}(?:[.,]\d{1,3})?)\s*(ml|l|lt|kg|g|gr)\b", re.I)


def sem_acento(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c)
    )


def normalizar(texto: str) -> str:
    """Maiusculas, sem acento, sem pontuacao, espacos colapsados."""
    t = sem_acento(str(texto or "")).upper()
    t = re.sub(r"[^\w\s]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def normalizar_medida(texto: str) -> str:
    """
    Como normalizar(), mas PRESERVA virgula e ponto -- necessario porque
    'REFRIGERANTE 1,5L' viraria '1 5L' e a medida seria lida como 5 L.
    """
    t = sem_acento(str(texto or "")).upper()
    t = re.sub(r"[^\w\s,.]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def termo_de_busca(descricao: str) -> str:
    """
    Nucleo do produto para enviar ao portal: remove palavras de embalagem e a
    propria medida (o portal casa melhor sem ela), preservando marca.
    """
    base = normalizar_medida(descricao)
    base = _RE_MULTI.sub(" ", base)
    base = _RE_MEDIDA.sub(" ", base)
    base = re.sub(r"[^\w\s]", " ", base)
    palavras = [
        p for p in base.split()
        if p.lower() not in RUIDO and not re.fullmatch(r"[\d.,]+", p)
    ]
    return " ".join(palavras[:6]).strip() or normalizar(descricao)


@dataclass
class Medida:
    quantidade: float | None = None   # na unidade base
    base: str | None = None           # "ml" | "g" | None
    embalagens: int = 1               # multipack
    ambigua: bool = False             # descricao permite duas leituras

    @property
    def total(self) -> float | None:
        if self.quantidade is None:
            return None
        return self.quantidade * self.embalagens


def extrair_medida(descricao: str) -> Medida:
    """
    Extrai volume/peso e converte para unidade base (ml ou g).
    '1L' -> 1000 ml; '500ML' -> 500 ml; '12x500ml' -> 500 ml x 12 embalagens.

    Descricao AMBIGUA -- caso real do Busca Preco AM: "DETERGENTE GUAMA MACA
    500ML - 24X500ML" a R$ 1,99. Lido como fardo, sao 12 litros por R$ 1,99, um
    R$/litro absurdo que inventa 82% de economia. O preco da NFC-e e da unidade
    vendida, e a mencao "500ML" ANTES do multipack mostra que o produto e a
    unidade; "24X500ML" so descreve a caixa de origem.

    Regra: se a mesma medida aparece solta FORA do trecho do multipack, vale a
    leitura de UNIDADE (nunca infla o volume, entao nunca cria economia falsa) e
    a medida sai marcada `ambigua` para o relatorio poder ressalvar.
    """
    texto = normalizar_medida(descricao)

    m = _RE_MULTI.search(texto)
    if m:
        n = int(m.group(1))
        valor = float(m.group(2).replace(",", "."))
        base, fator = UNIDADES[m.group(3).lower()]
        unitaria = valor * fator

        # a mesma medida aparece fora do trecho "24X500ML"?
        fora = texto[:m.start()] + " " + texto[m.end():]
        for outra in _RE_MEDIDA.finditer(fora):
            v2 = float(outra.group(1).replace(",", "."))
            b2, f2 = UNIDADES[outra.group(2).lower()]
            if b2 == base and abs(v2 * f2 - unitaria) < 1e-9:
                return Medida(unitaria, base, 1, ambigua=True)
        return Medida(unitaria, base, n)

    m = _RE_MEDIDA.search(texto)
    if m:
        valor = float(m.group(1).replace(",", "."))
        base, fator = UNIDADES[m.group(2).lower()]
        return Medida(valor * fator, base, 1)

    return Medida()


def descartar_outliers(valores: list[float | None], piso: float = 0.15,
                       minimo_amostra: int = 4) -> tuple[list[int], float | None]:
    """
    Indices de valores absurdamente baixos frente a mediana dos demais.

    Existe porque o portal do AM devolveu uma NFC-e real de "REFRIGERANTE COCA
    COLA LT 350ML" a **R$ 0,01** num restaurante -- brinde, cortesia ou ajuste
    fiscal. E preco verdadeiro na nota e mentira como preco de mercado: sozinho,
    virava "99,43% de economia" no relatorio.

    Nao e filtro de preco baixo (isso esconderia economia real): e filtro de
    valor que nao pertence a mesma distribuicao das outras ofertas do MESMO
    produto. Com menos de `minimo_amostra` ofertas a mediana nao diz nada e nada
    e descartado.
    """
    import statistics
    limpos = [v for v in valores if v is not None and v > 0]
    if len(limpos) < minimo_amostra:
        return [], None
    mediana = statistics.median(limpos)
    if mediana <= 0:
        return [], None
    corte = mediana * piso
    fora = [i for i, v in enumerate(valores) if v is not None and 0 < v < corte]
    return fora, mediana


def cobertura_tokens(desc_planilha: str, desc_portal: str) -> float:
    """
    Fracao das palavras do produto da planilha que aparecem na descricao do
    portal. Cobertura ZERO significa que o portal devolveu outra coisa: e o caso
    de "CEBOLA", que no Busca Preco AM casa por substring com "BOLA DE ISOPOR".
    """
    tokens_a = set(termo_de_busca(desc_planilha).split())
    tokens_b = set(normalizar(desc_portal).split())
    if not tokens_a:
        return 0.0
    return len(tokens_a & tokens_b) / len(tokens_a)


def preco_por_base(preco: float | None, medida: Medida) -> float | None:
    """
    Preco por unidade base (R$/ml ou R$/g). E ESTE o numero que deve ser
    comparado: 1L a R$1,50 e 500ml a R$0,75 tem o MESMO preco por litro.
    """
    if preco is None or medida.total in (None, 0):
        return None
    return float(preco) / float(medida.total)


def comparaveis(a: Medida, b: Medida, tolerancia: float = 0.02) -> bool:
    """Duas medidas sao comparaveis se tem a mesma base e total equivalente."""
    if a.total is None or b.total is None:
        return False
    if a.base != b.base:
        return False
    maior = max(a.total, b.total)
    return abs(a.total - b.total) / maior <= tolerancia


def nivel_confianca(
    desc_planilha: str, desc_portal: str, gtin_igual: bool = False
) -> str:
    """ALTA | MEDIA | BAIXA -- o quanto confiar no casamento por descricao."""
    if gtin_igual:
        return "ALTA"

    m_a, m_b = extrair_medida(desc_planilha), extrair_medida(desc_portal)
    cobertura = cobertura_tokens(desc_planilha, desc_portal)

    # Nenhuma palavra em comum: o portal casou por substring e devolveu outro
    # produto ("CEBOLA" -> "BOLA DE ISOPOR"). Nao e um casamento fraco, e um
    # nao-casamento; entra como RUIDO e nunca vira preco de referencia.
    if cobertura == 0:
        return "RUIDO"

    # Cobertura sozinha engana quando o produto tem poucas palavras: "CEBOLA"
    # cobre 100% de "SALGADINHO AKIMILHO 30 GR CEBOLA E SALSA" e o salgadinho de
    # R$ 0,89 viraria "cebola barata". O Dice olha os DOIS lados: quanto do termo
    # foi coberto e quanto da descricao do portal e outra coisa.
    tokens_a = set(termo_de_busca(desc_planilha).split())
    tokens_b = set(normalizar(desc_portal).split())
    dice = 2 * len(tokens_a & tokens_b) / (len(tokens_a) + len(tokens_b)) if tokens_b else 0.0

    if dice < 0.35:
        # a descricao do portal e majoritariamente outro produto. So sobrevive
        # como casamento fraco se ao menos a embalagem bater.
        return "BAIXA" if comparaveis(m_a, m_b) else "RUIDO"

    if comparaveis(m_a, m_b) and cobertura >= 0.6:
        return "ALTA"
    if comparaveis(m_a, m_b) or cobertura >= 0.6:
        return "MEDIA"
    return "BAIXA"


def _primeiro(d: dict, *chaves: str) -> Any:
    """
    Primeiro valor nao-nulo entre `chaves`. Diferente de `a or b`, aceita 0 e ""
    como valores legitimos -- so descarta None e chave ausente.
    """
    if not isinstance(d, dict):
        return None
    for c in chaves:
        if d.get(c) is not None:
            return d[c]
    return None


def parse_preco(texto: Any) -> float | None:
    """'R$ 1.234,56' -> 1234.56 ; aceita float/int direto."""
    if texto is None or isinstance(texto, bool):
        return None
    if isinstance(texto, (int, float)):
        return float(texto)
    t = re.sub(r"[^\d,.\-]", "", str(texto))
    if not t:
        return None
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")   # 1.234,56
    elif "," in t:
        t = t.replace(",", ".")                    # 1,50
    try:
        return float(t)
    except ValueError:
        return None


# --------------------------------------------------------------------------
# 2. Geografia
# --------------------------------------------------------------------------

def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Distancia em km, em LINHA RETA (nao rota rodoviaria nem fluvial)."""
    r = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = p2 - p1
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def _cache_load() -> dict:
    try:
        with open(CACHE_GEO, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _cache_save(d: dict) -> None:
    try:
        os.makedirs(os.path.dirname(CACHE_GEO), exist_ok=True)
        with open(CACHE_GEO, "w", encoding="utf-8") as fh:
            json.dump(d, fh)
    except Exception:
        pass


def geocodificar(municipio: str, uf: str, session=None) -> tuple[float, float] | None:
    """
    Resolve municipio -> (lat, lon) via Nominatim/OSM, com cache em disco.
    Devolve None se nao conseguir. NUNCA inventa coordenada: sem geocodificacao,
    a distancia sai como 'n/d' no relatorio.
    """
    chave = f"{normalizar(municipio)}|{uf.upper()}"
    cache = _cache_load()
    if chave in cache:
        v = cache[chave]
        return (v[0], v[1]) if v else None

    try:
        import requests
        s = session or requests.Session()
        resp = s.get(
            "https://nominatim.openstreetmap.org/search",
            params={
                "q": f"{municipio}, {uf}, Brasil",
                "format": "json",
                "limit": 1,
                "countrycodes": "br",
            },
            headers={"User-Agent": UA},
            timeout=20,
        )
        resp.raise_for_status()
        dados = resp.json()
        coord = (float(dados[0]["lat"]), float(dados[0]["lon"])) if dados else None
    except Exception:
        coord = None

    cache[chave] = list(coord) if coord else None
    _cache_save(cache)
    time.sleep(1.1)          # politica de uso do Nominatim: 1 req/s
    return coord


# --------------------------------------------------------------------------
# 3. Resultado normalizado
# --------------------------------------------------------------------------

@dataclass
class Oferta:
    descricao: str
    preco: float
    estabelecimento: str = ""
    endereco: str = ""
    municipio: str = ""
    data_venda: str = ""
    gtin: str = ""
    distancia_km: float | None = None
    latitude: float | None = None
    longitude: float | None = None
    fonte: str = ""
    medida: Medida = field(default_factory=Medida)

    def to_row(self) -> dict:
        d = asdict(self)
        d["medida_total"] = self.medida.total
        d["medida_base"] = self.medida.base
        d.pop("medida", None)
        d["preco_por_base"] = preco_por_base(self.preco, self.medida)
        return d


# --------------------------------------------------------------------------
# 4. Adaptador Amazonas -- Busca Preco AM
# --------------------------------------------------------------------------

class AdapterAM:
    """
    Contrato observado (derivado de cliente publico existente):

        GET  https://buscapreco.sefaz.am.gov.br/home        -> cookie JSESSIONID
        POST https://buscapreco.sefaz.am.gov.br/item/grupo/page/<n>
             corpo multipart: descricaoProd=<termo>
             resposta: HTML em latin-1

    Cards no HTML:
        div.card.small.p.hoverable
          .card-title b                       nome do produto
          .tb-valor-25 b                      preco
          .tb-valor-10                        data/hora da venda
          .truncate.tooltipped.padding10 b    estabelecimento
          .valign-wrapper div span            endereco

    Janela de dados: 48h por padrao. Sem login e sem captcha.

    O formulario #frmConsulta tem, ALEM de descricaoProd: cdGtin (busca exata
    por codigo de barras), tipoConsulta (0/24/48/168 horas), distancia
    (2/5/10/9999 km), municipio, precoMinimo e precoMaximo. Este adaptador NAO
    usa os filtros de raio e municipio -- enviei distancia=2 com coordenada de
    Manaus em 09/09/2026 e o resultado nao mudou, entao nao afirmo que
    funcionam por POST. O municipio de cada oferta sai do endereco do
    estabelecimento, e a distancia, da coordenada que o card traz.
    """

    BASE = "https://buscapreco.sefaz.am.gov.br"
    HOME = f"{BASE}/home"
    BUSCA = f"{BASE}/item/grupo/p"
    # Rota real usada pelo formulario da home (capturada no navegador em
    # 08/09/2026): POST /item/grupo/p/<n>. A antiga /item/grupo/page/<n> ainda
    # responde igual, mas /p/ e a que o portal usa de fato.
    CANDIDATOS = ["/item/grupo/p", "/item/grupo/page", "/item/page"]

    def __init__(self, pausa: float = 1.5, paginas: int = 2):
        import requests
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": UA})
        self.pausa = pausa
        self.paginas = paginas
        self._rota = None
        self._raw: list[str] = []

    def abrir_sessao(self) -> None:
        r = self.s.get(self.HOME, timeout=30)
        r.raise_for_status()
        if "JSESSIONID" not in self.s.cookies.get_dict():
            # Alguns servidores so emitem o cookie apos a primeira busca.
            pass

    def _post(self, rota: str, pagina: int, termo: str):
        """
        Corpo urlencoded EXATAMENTE como o formulario da home envia (capturado
        no navegador em 08/09/2026):

            POST /item/grupo/p/1
            Content-Type: application/x-www-form-urlencoded
            descricaoProd=cebola&latitude=&longitude=&g-recaptcha-response=&action=

        NAO usar multipart: o servidor nao faz o bind do campo, devolve zero
        resultados e a pagina exibe a BUSCA ANTERIOR guardada na sessao. Como o
        HTML sempre contem a palavra "card" (e nome de classe CSS), um teste por
        substring nao percebe isso -- e o que fazia termos diferentes voltarem o
        mesmo produto. Por isso a checagem aqui e por card de verdade.
        """
        url = f"{self.BASE}{rota}/{pagina}"
        r = self.s.post(url, data={
            "descricaoProd": termo, "latitude": "", "longitude": "",
            "g-recaptcha-response": "", "action": "",
        }, timeout=45)
        r.encoding = "latin-1"
        return r

    @staticmethod
    def _tem_cards(html: str) -> bool:
        """Ha cards de resultado de verdade? (nao basta a substring 'card')"""
        from bs4 import BeautifulSoup
        sopa = BeautifulSoup(html, "lxml")
        return bool(sopa.select("div.card.small.p.hoverable") or
                    sopa.select("div.card .card-title b"))

    def _descobrir_rota(self, termo: str) -> str | None:
        for rota in self.CANDIDATOS:
            try:
                r = self._post(rota, 1, termo)
            except Exception:
                continue
            if r.status_code < 400 and self._tem_cards(r.text):
                return rota
        return None

    def buscar(self, termo: str) -> list[Oferta]:
        from bs4 import BeautifulSoup

        if self._rota is None:
            self.abrir_sessao()
            self._rota = self._descobrir_rota(termo)
            if self._rota is None:
                raise RuntimeError(
                    "Nao localizei a rota de busca do Busca Preco AM. O portal "
                    "provavelmente foi reescrito: abra buscapreco.sefaz.am.gov.br "
                    "num navegador, faca uma busca e capture a requisicao real na "
                    "aba de rede; depois atualize AdapterAM.CANDIDATOS."
                )

        ofertas: list[Oferta] = []
        for pagina in range(1, self.paginas + 1):
            r = self._post(self._rota, pagina, termo)
            self._raw.append(r.text)
            if r.status_code >= 400:
                break
            sopa = BeautifulSoup(r.text, "lxml")
            cards = sopa.select("div.card.small.p.hoverable") or sopa.select("div.card")
            if not cards:
                break
            for card in cards:
                nome = self._txt(card, ".card-title b")
                preco = parse_preco(self._txt(card, ".tb-valor-25 b"))
                if not nome or preco is None:
                    continue
                endereco = self._txt(card, ".valign-wrapper div span")
                lat, lon = self._coordenadas(card)
                ofertas.append(
                    Oferta(
                        descricao=nome,
                        preco=preco,
                        gtin=self._gtin(card),
                        estabelecimento=self._txt(card, ".truncate.tooltipped.padding10 b"),
                        endereco=endereco,
                        municipio=self._municipio(endereco),
                        data_venda=self._data_venda(self._txt(card, ".tb-valor-10")),
                        latitude=lat,
                        longitude=lon,
                        fonte="Busca Preco AM",
                        medida=extrair_medida(nome),
                    )
                )
            time.sleep(self.pausa)
        return ofertas

    @staticmethod
    def _txt(no, seletor: str) -> str:
        achado = no.select_one(seletor)
        return achado.get_text(" ", strip=True) if achado else ""

    @staticmethod
    def _data_venda(relativo: str) -> str:
        """
        O Busca Preco AM informa a venda de forma RELATIVA -- "Ha 7 hora(s) 48
        minuto(s) 36 segundo(s)", as vezes com "N dia(s)" na frente. Guardado
        assim, o relatorio mente ao ser lido no dia seguinte. Converte para data
        absoluta no fuso de Manaus (UTC-4) e mantem o relativo entre parenteses:

            "Ha 7 hora(s) 48 minuto(s)" -> "08/09/2026 09:12 (ha 7h48)"

        Texto fora do padrao volta inalterado -- nunca inventa data.
        """
        import datetime as _dt

        if not relativo:
            return ""
        texto = sem_acento(relativo).lower()
        unidades = {"dia": 0, "hora": 0, "minuto": 0}
        for chave in unidades:
            m = re.search(r"(\d+)\s*" + chave, texto)
            if m:
                unidades[chave] = int(m.group(1))
        if not any(unidades.values()) and "segundo" not in texto:
            return relativo          # nao reconheci: devolve o que o portal disse

        agora = _dt.datetime.now(_dt.timezone(_dt.timedelta(hours=-4)))   # Manaus
        quando = agora - _dt.timedelta(
            days=unidades["dia"], hours=unidades["hora"], minutes=unidades["minuto"]
        )
        if unidades["dia"]:
            rel = "ha %dd %02dh%02d" % (unidades["dia"], unidades["hora"], unidades["minuto"])
        else:
            rel = "ha %dh%02d" % (unidades["hora"], unidades["minuto"])
        return quando.strftime("%d/%m/%Y %H:%M") + " (" + rel + ")"

    @staticmethod
    def _gtin(card) -> str:
        """
        Codigo de barras do produto. O card nao o mostra na tela, mas o traz no
        gatilho do proprio portal:  onclick="findByGtin(7893500018469);"

        E o codigo com que se refaz a busca exata deste item no portal, sem
        depender de descricao.

        SO o gatilho vale. Nao procure digitos soltos no card: a coordenada do
        mapa -- refreshMap(-59.99, -3.05139623399998) -- tem 14 digitos na parte
        decimal e passava por "EAN-14", produzindo um codigo que manda o usuario
        buscar outro produto. Sem findByGtin, devolve vazio, que e honesto.
        """
        m = re.search(r"findByGtin\(\s*(\d{8,14})\s*\)", str(card))
        return m.group(1) if m else ""

    @staticmethod
    def _coordenadas(card) -> tuple[float | None, float | None]:
        """
        Coordenada do estabelecimento, tambem escondida num gatilho:
            onclick="javascript:refreshMap( -59.9931023, -3.0354544);"
        A ordem e (LONGITUDE, LATITUDE) -- confirmado por Manaus, que fica em
        lat -3, lon -60. Devolve (lat, lon); (None, None) se nao houver.
        """
        m = re.search(r"refreshMap\(\s*(-?\d+\.?\d*)\s*,\s*(-?\d+\.?\d*)\s*\)", str(card))
        if not m:
            return None, None
        lon, lat = float(m.group(1)), float(m.group(2))
        if not (-90 <= lat <= 90 and -180 <= lon <= 180):
            return None, None
        return lat, lon

    @staticmethod
    def _municipio(endereco: str) -> str:
        """Extrai o municipio de um endereco livre. Heuristica -- pode falhar."""
        if not endereco:
            return ""
        partes = [p.strip() for p in re.split(r"[-,/]", endereco) if p.strip()]
        for parte in reversed(partes):
            limpo = re.sub(r"\b(AM|CEP|\d{5}-?\d{3})\b", "", parte, flags=re.I).strip()
            if limpo and not limpo.isdigit() and len(limpo) > 2:
                return limpo.title()
        return ""


# --------------------------------------------------------------------------
# 5. Adaptador Preco da Hora (PB e BA -- mesmo codigo-base)
# --------------------------------------------------------------------------

class AdapterPrecoDaHora:
    """
    Contrato observado (derivado de clientes publicos existentes):

        GET  <base>/            -> Set-Cookie + CSRF token no atributo
                                   data-id do elemento #validate
        POST <base>/sugestao/   corpo: item=<termo>
                                -> {"resultado": [{"gtin": "...", ...}]}
        POST <base>/produtos/   corpo: gtin, latitude, longitude, raio, horas,
                                   precomin, precomax, ordenar=preco.asc,
                                   pagina, processo=carregar, totalRegistros,
                                   totalPaginas, pageview=lista
                                -> {"resultado": [{"produto": {...},
                                                   "estabelecimento": {...}}],
                                    "totalRegistros", "totalPaginas", "dias", ...}

        Cada registro do `resultado` e ANINHADO (verificado em 08/09/2026):
        preco em produto.precoUnitario, nome em produto.descricao, data da
        venda em produto.data, e o vendedor em estabelecimento.* --
        nomeEstabelecimento, endLogradouro, endNumero, bairro, municipio,
        latitude, longitude, distancia. Ver AdapterPrecoDaHora._ler_item.

        O campo `dias` da resposta e apenas o eco de `horas`, com nome
        infeliz: enviando horas=3 a resposta traz dias=3 e a janela observada
        e de ~3 HORAS. Enviar `dias` no corpo nao tem efeito.

        Headers obrigatorios no POST:
            Content-Type: application/x-www-form-urlencoded; charset=UTF-8
            X-CSRFToken: <token>
            Referer: <base>/

    Raio padrao 10 km, maximo 30 km. Defasagem de ate 3 dias (parametro
    `horas`, ex. 72). O fluxo descricao -> /sugestao/ -> gtin -> /produtos/ e
    exatamente o que permite casar por descricao sem exigir codigo de barras.

    ATENCAO PB: precodahora.tcepb.tc.br fica atras de verificacao anti-bot do
    Cloudflare. Cliente HTTP simples recebe 403. Rode de um navegador real ou
    valide o fluxo no portal da Bahia, que nao tem anti-bot.
    """

    BASES = {
        "PB": "https://precodahora.tcepb.tc.br",
        "BA": "https://precodahora.ba.gov.br",
    }

    def __init__(self, uf: str, lat: float, lon: float, raio_km: int = 30,
                 horas: int = 72, pausa: float = 1.5, paginas: int = 2):
        import requests
        uf = uf.upper()
        if uf not in self.BASES:
            raise ValueError(f"Preco da Hora nao cobre {uf}")
        self.base = self.BASES[uf]
        self.uf = uf
        self.lat, self.lon = lat, lon
        self.raio = min(int(raio_km), 30)      # o portal nao aceita mais que 30
        self.horas = horas
        self.pausa = pausa
        self.paginas = paginas
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": UA})
        self.csrf = None
        self._raw: list[str] = []

    def abrir_sessao(self) -> None:
        from bs4 import BeautifulSoup
        r = self.s.get(self.base + "/", timeout=40)
        if r.status_code == 403:
            raise RuntimeError(
                f"{self.base} respondeu 403 -- verificacao anti-bot do Cloudflare. "
                "Este host exige navegador real; nao ha como contornar por HTTP. "
                "Para validar o fluxo, use --uf BA (mesmo codigo-base, sem anti-bot)."
            )
        r.raise_for_status()
        no = BeautifulSoup(r.text, "lxml").select_one("#validate")
        self.csrf = no.get("data-id") if no else None
        if not self.csrf:
            raise RuntimeError(
                "Nao encontrei o CSRF token (#validate[data-id]) na home. "
                "O portal pode ter mudado o mecanismo."
            )

    def _post(self, rota: str, corpo: dict) -> dict:
        cabecalhos = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-CSRFToken": self.csrf or "",
            "Referer": self.base + "/",
            "X-Requested-With": "XMLHttpRequest",
        }
        r = self.s.post(self.base + rota, data=corpo, headers=cabecalhos, timeout=45)
        self._raw.append(r.text[:20000])
        r.raise_for_status()
        try:
            return r.json()
        except Exception:
            return {}

    def sugerir_gtins(self, termo: str, limite: int = 3) -> list[dict]:
        dados = self._post("/sugestao/", {"item": termo})
        return (dados.get("resultado") or [])[:limite]

    def buscar(self, termo: str) -> list[Oferta]:
        if self.csrf is None:
            self.abrir_sessao()

        ofertas: list[Oferta] = []
        for sugestao in self.sugerir_gtins(termo):
            gtin = str(sugestao.get("gtin") or "").strip()
            if not gtin:
                continue
            for pagina in range(1, self.paginas + 1):
                dados = self._post("/produtos/", {
                    "gtin": gtin,
                    "latitude": self.lat,
                    "longitude": self.lon,
                    "raio": self.raio,
                    "horas": self.horas,
                    "precomin": 0,
                    "precomax": 0,
                    "ordenar": "preco.asc",
                    "pagina": pagina,
                    "processo": "carregar",
                    "totalRegistros": 0,
                    "totalPaginas": 0,
                    "pageview": "lista",
                })
                itens = dados.get("resultado") or []
                if not itens:
                    break
                for item in itens:
                    oferta = self._ler_item(item, gtin, termo)
                    if oferta is not None:
                        ofertas.append(oferta)
                time.sleep(self.pausa)
        return ofertas

    def _ler_item(self, item: dict, gtin: str, termo: str) -> Oferta | None:
        """
        Converte um registro de /produtos/ em Oferta.

        Contrato ATUAL (capturado em 08/09/2026 no portal da BA): cada registro
        vem ANINHADO em duas chaves, `produto` e `estabelecimento` --

            {"produto": {"gtin", "descricao", "precoUnitario", "precoLiquido",
                         "unidade", "data", ...},
             "estabelecimento": {"nomeEstabelecimento", "endLogradouro",
                         "endNumero", "bairro", "municipio", "uf",
                         "latitude", "longitude", "distancia", ...}}

        O contrato antigo (plano, com `valor`/`nm_emp`/`nm_mun`) ainda e aceito
        como alternativa, para o caso de o portal reverter.
        """
        prod = item.get("produto") if isinstance(item.get("produto"), dict) else None
        est = item.get("estabelecimento") if isinstance(item.get("estabelecimento"), dict) else {}
        p = prod if prod is not None else item      # plano = o proprio registro

        preco = parse_preco(_primeiro(p, "precoUnitario", "precoLiquido", "valor",
                                      "preco", "valorProduto"))
        if preco is None:
            return None

        nome = str(_primeiro(p, "descricao", "desc") or termo)
        endereco = " ".join(str(x) for x in (
            est.get("endLogradouro") or est.get("nm_logr") or est.get("endereco") or "",
            est.get("endNumero") or "",
            est.get("bairro") or "",
        ) if str(x).strip()).strip()

        # Distancia: calculada aqui por haversine a partir da coordenada do
        # estabelecimento (verificavel), com a do portal como alternativa.
        lat, lon = parse_preco(est.get("latitude")), parse_preco(est.get("longitude"))
        if lat is not None and lon is not None:
            dist = round(haversine(self.lat, self.lon, lat, lon), 1)
        else:
            dist = parse_preco(_primeiro(est, "distancia") or _primeiro(item, "distkm", "distancia"))

        return Oferta(
            descricao=nome,
            preco=preco,
            gtin=str(_primeiro(p, "gtin", "codProduto") or gtin),
            estabelecimento=str(_primeiro(est, "nomeEstabelecimento", "nm_emp", "nome") or ""),
            endereco=endereco,
            municipio=str(_primeiro(est, "municipio", "nm_mun") or ""),
            data_venda=str(_primeiro(p, "data", "datahora") or ""),
            distancia_km=dist,
            fonte=f"Preco da Hora {self.uf}",
            medida=extrair_medida(nome),
        )


class AdapterPrecoDaHoraNavegador(AdapterPrecoDaHora):
    """
    Mesmo contrato do Preco da Hora, mas as chamadas saem de um navegador real
    (Chrome via Playwright), que e o que a Paraiba exige.

    Em 08/09/2026 https://precodahora.tcepb.tc.br devolve 403 para qualquer
    cliente HTTP e redireciona o navegador para /cf_captcha -- um desafio do
    Cloudflare. Esta classe NAO resolve o desafio: ela abre a janela e ESPERA
    ate `espera_humano` segundos para que uma PESSOA conclua a verificacao. Feito
    isso, a sessao do navegador ja esta valida e as consultas seguem pelas mesmas
    rotas /sugestao/ e /produtos/, executadas no contexto da pagina.

    Sem alguem para resolver o desafio, esta classe falha com uma mensagem
    explicita. Nao ha, e nao deve haver, caminho automatico para isso.
    """

    def __init__(self, uf: str, lat: float, lon: float, raio_km: int = 30,
                 horas: int = 72, pausa: float = 1.5, paginas: int = 2,
                 espera_humano: int = 180, visivel: bool = True,
                 perfil: str | None = None):
        super().__init__(uf, lat, lon, raio_km, horas, pausa, paginas)
        self.espera_humano = espera_humano
        self.visivel = visivel
        # Perfil persistente: o desafio do Cloudflare, resolvido UMA vez pela
        # pessoa, deixa cookie aqui e as coletas seguintes nao pedem de novo.
        self.perfil = perfil
        self._pw = self._nav = self._pag = None

    def abrir_sessao(self) -> None:
        from playwright.sync_api import sync_playwright

        self._pw = sync_playwright().start()
        if self.perfil:
            os.makedirs(self.perfil, exist_ok=True)
            self._nav = self._pw.chromium.launch_persistent_context(
                self.perfil, channel="chrome", headless=not self.visivel,
                viewport={"width": 1280, "height": 800},
                locale="pt-BR", timezone_id="America/Fortaleza")
            self._pag = self._nav.pages[0] if self._nav.pages else self._nav.new_page()
        else:
            self._nav = self._pw.chromium.launch(channel="chrome",
                                                 headless=not self.visivel)
            self._pag = self._nav.new_page()
        self._pag.goto(self.base + "/", wait_until="domcontentloaded", timeout=90000)

        # Espera ativa pelo fim da verificacao: quem resolve e o humano na janela.
        #
        # Tudo aqui e a prova de navegacao. Quando a pessoa clica no checkbox, o
        # Cloudflare navega da /cf_captcha para a home -- e um evaluate em curso
        # nesse instante morre com "Execution context was destroyed". Isso NAO e
        # falha da verificacao: e o sucesso dela. Cada tentativa vive em try, e o
        # laco simplesmente tenta de novo na volta seguinte.
        limite = time.time() + self.espera_humano
        while time.time() < limite:
            try:
                self._pag.wait_for_load_state("domcontentloaded", timeout=5000)
            except Exception:
                pass
            try:
                csrf = self._pag.evaluate(
                    "() => { const n = document.querySelector('#validate');"
                    " return n ? n.getAttribute('data-id') : null; }")
            except Exception:
                csrf = None              # navegando agora; reavalia no proximo ciclo
            if csrf:
                self.csrf = csrf
                return
            try:
                self._pag.wait_for_timeout(3000)
            except Exception:
                time.sleep(3)

        try:
            self._pag.screenshot(path="_pb_verificacao.png")
            dica = " Uma foto da tela ficou em _pb_verificacao.png."
        except Exception:
            dica = ""
        raise RuntimeError(
            "O portal da Paraiba continua na verificacao do Cloudflare "
            f"({self._pag.url}). Ha um checkbox 'Confirme que e humano' na janela do "
            "Chrome que foi aberta: UMA PESSOA precisa clicar nele -- por acesso "
            "remoto a esta maquina, se for o caso. Depois disso a coleta segue "
            "sozinha, e com --perfil o clique vale para as proximas execucoes."
            + dica
        )

    def _post(self, rota: str, corpo: dict) -> dict:
        """Mesma requisicao, emitida de dentro da pagina (herda cookies e sessao)."""
        corpo_txt = "&".join(f"{k}={requests_quote(str(v))}" for k, v in corpo.items())
        r = self._pag.evaluate("""async ([base, rota, csrf, corpo]) => {
            const r = await fetch(base + rota, {
                method: 'POST',
                headers: {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                          'X-CSRFToken': csrf, 'X-Requested-With': 'XMLHttpRequest'},
                body: corpo, credentials: 'include'});
            return {status: r.status, texto: await r.text()};
        }""", [self.base, rota, self.csrf or "", corpo_txt])
        self._raw.append(r["texto"][:20000])
        if r["status"] >= 400:
            raise RuntimeError(f"{rota} respondeu {r['status']} mesmo no navegador")
        try:
            return json.loads(r["texto"])
        except Exception:
            return {}

    def fechar(self) -> None:
        for obj, metodo in ((self._nav, "close"), (self._pw, "stop")):
            try:
                if obj:
                    getattr(obj, metodo)()
            except Exception:
                pass


def requests_quote(v: str) -> str:
    from urllib.parse import quote
    return quote(v, safe="")


class AdapterColetado:
    """
    Nao consulta nada: recebe as respostas CRUAS que outra pessoa (ou o agente,
    pela extensao do Chrome) ja colheu do portal, e as converte em Oferta com o
    mesmo parser usado na coleta por HTTP.

    Existe para o caminho da Paraiba: o portal fica atras de um CAPTCHA do
    Cloudflare, entao quem coleta e o navegador do proprio usuario, na aba dele,
    depois que ELE resolve o desafio. Toda a logica de casamento, unidade base,
    outlier e ressalva continua aqui no Python -- o navegador e so transporte.

    Formato do arquivo (`--ofertas-json`): termo consultado -> lista de respostas
    cruas de /produtos/, como o portal devolveu, sem edicao:

        {"DETERGENTE YPE": [ {"resultado": [{"produto": {...},
                                             "estabelecimento": {...}}, ...]} ]}

    Aceita tambem a lista de registros direto, sem o envelope {"resultado": ...}.
    """

    def __init__(self, uf: str, lat: float, lon: float, coletado: dict):
        self.uf = uf.upper()
        self.lat, self.lon = lat, lon
        self.coletado = coletado or {}
        self._raw: list[str] = []
        # reaproveita _ler_item sem abrir sessao nem tocar na rede
        self._leitor = AdapterPrecoDaHora.__new__(AdapterPrecoDaHora)
        self._leitor.uf = self.uf
        self._leitor.lat, self._leitor.lon = lat, lon

    def buscar(self, termo: str) -> list[Oferta]:
        blocos = self.coletado.get(termo)
        if blocos is None:
            # tolera diferenca de acentuacao/caixa entre o termo pedido e o gravado
            alvo = normalizar(termo)
            for chave, valor in self.coletado.items():
                if chave.startswith("_"):
                    continue                 # metadados da coleta
                if normalizar(chave) == alvo:
                    blocos = valor
                    break
        if not blocos:
            return []

        # Formato do Preco da Hora PB reescrito em Next.js (setembro/2026): a
        # coleta vem da pagina do produto, ja com o nome e o GTIN resolvidos, e
        # cada loja e uma tupla [loja, preco, data, bairro, cidade].
        if isinstance(blocos, dict) and "lojas" in blocos:
            return self._ler_pagina_produto(blocos)

        if isinstance(blocos, dict):
            blocos = [blocos]

        return self._ler_blocos(blocos, termo)

    def _ler_pagina_produto(self, bloco: dict) -> list[Oferta]:
        """
        Coleta da pagina de produto do Preco da Hora PB (Next.js).

        O nome e o GTIN sao do PRODUTO (a pagina e de um item so), e cada loja
        traz preco, data e localizacao. Diferente do AM, o preco aqui e a
        **media das ultimas vendas daquele lojista** -- o proprio portal diz
        isso -- e nao o valor de uma nota individual. O relatorio precisa
        declarar essa diferenca; ela vai em `fonte`.
        """
        nome = str(bloco.get("produto") or "")
        gtin = str(bloco.get("gtin") or "")
        medida = extrair_medida(nome)
        ofertas: list[Oferta] = []
        for linha in bloco.get("lojas") or []:
            if isinstance(linha, dict):
                loja, preco = linha.get("loja"), linha.get("preco")
                data, bairro = linha.get("data"), linha.get("bairro")
                cidade = linha.get("cidade")
            elif isinstance(linha, (list, tuple)) and len(linha) >= 2:
                loja, preco = linha[0], linha[1]
                data = linha[2] if len(linha) > 2 else ""
                bairro = linha[3] if len(linha) > 3 else ""
                cidade = linha[4] if len(linha) > 4 else ""
            else:
                continue
            valor = parse_preco(preco)
            if valor is None or not loja:
                continue
            ofertas.append(Oferta(
                descricao=nome, preco=valor, gtin=gtin,
                estabelecimento=str(loja),
                endereco=str(bairro or ""),
                municipio=str(cidade or ""),
                data_venda=str(data or ""),
                fonte=f"Preco da Hora {self.uf} (media do lojista)",
                medida=medida,
            ))
        self._raw.append(json.dumps(bloco, ensure_ascii=False)[:20000])
        return ofertas

    def _ler_blocos(self, blocos, termo: str) -> list[Oferta]:
        ofertas: list[Oferta] = []
        for bloco in blocos:
            registros = bloco.get("resultado") if isinstance(bloco, dict) else bloco
            for item in (registros or []):
                if not isinstance(item, dict):
                    continue
                gtin = str(((item.get("produto") or {}) if isinstance(item.get("produto"), dict)
                            else item).get("gtin") or "")
                oferta = self._leitor._ler_item(item, gtin, termo)
                if oferta is not None:
                    ofertas.append(oferta)
            self._raw.append(json.dumps(bloco, ensure_ascii=False)[:20000])
        return ofertas


# --------------------------------------------------------------------------
# 6. Planilha
# --------------------------------------------------------------------------

CAB_DESC = ("produto", "descricao", "item", "material", "mercadoria", "nome")
CAB_PRECO = ("preco", "valor", "custo", "preco atual", "precoatual", "vl unit", "vlunit")


def ler_planilha(caminho: str) -> tuple[list[dict], list[str]]:
    """Devolve (linhas como dict, cabecalhos)."""
    if caminho.lower().endswith((".csv", ".tsv", ".txt")):
        delim = "\t" if caminho.lower().endswith(".tsv") else None
        with open(caminho, encoding="utf-8-sig", newline="") as fh:
            amostra = fh.read(4096)
            fh.seek(0)
            if delim is None:
                delim = csv.Sniffer().sniff(amostra, delimiters=";,\t").delimiter
            leitor = csv.DictReader(fh, delimiter=delim)
            linhas = [dict(r) for r in leitor]
            return linhas, list(leitor.fieldnames or [])

    from openpyxl import load_workbook
    wb = load_workbook(caminho, data_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    cabecalhos = [str(c).strip() if c is not None else "" for c in next(it)]
    linhas = []
    for valores in it:
        if all(v is None or str(v).strip() == "" for v in valores):
            continue
        linhas.append({cabecalhos[i]: valores[i] for i in range(min(len(cabecalhos), len(valores)))})
    return linhas, cabecalhos


def detectar_colunas(cabecalhos: Iterable[str]) -> dict:
    achado = {"descricao": None, "preco": None, "gtin": None, "unidade": None}
    for cab in cabecalhos:
        n = normalizar(cab).lower()
        if achado["descricao"] is None and any(k in n for k in CAB_DESC):
            achado["descricao"] = cab
        if achado["preco"] is None and any(k in n for k in CAB_PRECO):
            achado["preco"] = cab
        if achado["gtin"] is None and any(k in n for k in ("gtin", "ean", "barras", "codigo de barras")):
            achado["gtin"] = cab
        if achado["unidade"] is None and any(k in n for k in ("unidade", "embalagem", "volume", "un")):
            achado["unidade"] = cab
    return achado


# --------------------------------------------------------------------------
# 7. Consolidacao
# --------------------------------------------------------------------------

def consolidar(descricao: str, preco_atual: float | None, ofertas: list[Oferta],
               municipio_base: str, uf: str, coord_base: tuple[float, float] | None,
               geocode: bool = True, max_alternativas: int = 5) -> dict:
    """
    Monta a linha de resultado. Compara por PRECO POR UNIDADE BASE quando as
    duas medidas sao conhecidas; cai para preco absoluto marcando confianca
    menor quando nao da.
    """
    medida_ref = extrair_medida(descricao)
    linha = {
        "descricao_planilha": descricao,
        "termo_consultado": termo_de_busca(descricao),
        "preco_atual": preco_atual,
        "medida_planilha": f"{medida_ref.total or ''} {medida_ref.base or ''}".strip(),
        "preco_atual_por_base": preco_por_base(preco_atual, medida_ref),
        "ofertas_encontradas": len(ofertas),
        "confianca_match": "NAO_ENCONTRADO",
        "menor_preco_municipio": None,
        "estabelecimento_municipio": "",
        "menor_preco_estado": None,
        "descricao_oferta": "",
        "gtin": "",
        "alternativas": [],
        "fornecedores_distintos": 0,
        "medida_oferta": "",
        "municipio_menor_preco": "",
        "estabelecimento_menor_preco": "",
        "endereco_menor_preco": "",
        "ofertas_descartadas_ruido": 0,
        "ofertas_descartadas_outlier": 0,
        "distancia_km": None,
        "data_venda": "",
        "economia_unitaria": None,
        "economia_percentual": None,
        "observacao": "",
    }
    _obs: list[str] = []

    def obs(*partes: str) -> None:
        """Acumula ressalvas: uma nunca apaga a outra."""
        texto = "".join(partes).strip()
        if texto and texto not in _obs:
            _obs.append(texto)

    if not ofertas:
        obs("nenhum resultado no portal para o termo consultado")
        linha["observacao"] = " | ".join(_obs)
        return linha

    for of in ofertas:
        of_conf = nivel_confianca(descricao, of.descricao, gtin_igual=False)
        setattr(of, "_conf", of_conf)
        if of.distancia_km is None and coord_base:
            if of.latitude is not None and of.longitude is not None:
                # coordenada do proprio estabelecimento: distancia exata, nao
                # a do centroide do municipio
                of.distancia_km = round(haversine(coord_base[0], coord_base[1],
                                                  of.latitude, of.longitude), 1)
            elif geocode and of.municipio:
                coord = geocodificar(of.municipio, uf)
                of.distancia_km = (
                    round(haversine(coord_base[0], coord_base[1], coord[0], coord[1]), 1)
                    if coord else None
                )

    def chave(of: Oferta):
        ppb = preco_por_base(of.preco, of.medida)
        return (ppb if ppb is not None else float("inf"), of.preco)

    # RUIDO nunca e utilizavel: e outro produto, nao um casamento fraco.
    ruidos = [o for o in ofertas if getattr(o, "_conf") == "RUIDO"]
    candidatas = [o for o in ofertas if getattr(o, "_conf") != "RUIDO"]
    linha["ofertas_descartadas_ruido"] = len(ruidos)
    if not candidatas:
        # o portal so devolveu outra coisa (casamento por substring). Nao ha
        # preco para este item: NAO_ENCONTRADO, e nenhum numero e reportado.
        obs(
            "nenhum resultado corresponde ao produto: o portal devolveu %d item(ns) "
            "sem nenhuma palavra em comum (ex.: %s) -- o portal casa por aproximacao "
            "de texto ('CEBOLA' puxa 'COLA' e 'BOLA'), e tudo foi descartado"
            % (len(ruidos), "; ".join(o.descricao[:32] for o in ruidos[:2]) or "-")
        )
        linha["observacao"] = " | ".join(_obs)
        return linha

    # BAIXA so entra se nao houver nada melhor -- e, quando entra, nao gera
    # economia: vira referencia sinalizada para conferencia humana.
    fortes = [o for o in candidatas if getattr(o, "_conf") != "BAIXA"]
    utilizaveis = fortes or candidatas
    so_baixa = not fortes

    # Preco que nao pertence a distribuicao do proprio produto (a lata de Coca a
    # R$ 0,01) sai antes de escolher o menor -- senao ele SEMPRE seria o menor.
    base_comp = [preco_por_base(o.preco, o.medida) or o.preco for o in utilizaveis]
    fora, mediana = descartar_outliers(base_comp)
    linha["ofertas_descartadas_outlier"] = len(fora)
    if fora and len(fora) < len(utilizaveis):
        descartadas = [utilizaveis[i] for i in fora]
        utilizaveis = [o for i, o in enumerate(utilizaveis) if i not in set(fora)]
        obs(
            "%d oferta(s) descartada(s) por preco fora da distribuicao do produto "
            "(ex.: %s a R$ %.2f, contra mediana equivalente das outras) -- NFC-e de "
            "brinde ou ajuste fiscal nao e preco de mercado"
            % (len(descartadas), descartadas[0].descricao[:34], descartadas[0].preco)
        )
    linha["confianca_match"] = max(
        (getattr(o, "_conf") for o in utilizaveis),
        key=lambda c: {"ALTA": 3, "MEDIA": 2, "BAIXA": 1}[c],
    )

    base_norm = normalizar(municipio_base)
    no_municipio = [o for o in utilizaveis if normalizar(o.municipio) == base_norm]
    if no_municipio:
        melhor_local = min(no_municipio, key=chave)
        linha["menor_preco_municipio"] = melhor_local.preco
        linha["estabelecimento_municipio"] = melhor_local.estabelecimento

    melhor = min(utilizaveis, key=chave)
    linha["menor_preco_estado"] = melhor.preco
    linha["municipio_menor_preco"] = melhor.municipio
    linha["estabelecimento_menor_preco"] = melhor.estabelecimento
    linha["endereco_menor_preco"] = melhor.endereco
    linha["distancia_km"] = melhor.distancia_km
    linha["data_venda"] = melhor.data_venda
    # Sem isto o numero fica indecifravel: o menor R$/litro de detergente e um
    # fardo "24 X 500ML." de R$ 58,00, que nao e o mesmo que comprar 1 litro.
    linha["descricao_oferta"] = melhor.descricao
    linha["medida_oferta"] = (f"{melhor.medida.total:g} {melhor.medida.base}"
                              if melhor.medida.total else "")
    linha["gtin"] = melhor.gtin

    # ALTERNATIVAS: uma so opcao nao serve para decidir -- o mais barato pode
    # estar longe, ser fardo ou ter vendido ha dois dias. Guarda as melhores por
    # preco na unidade base, uma por estabelecimento (a mesma loja repetida em
    # varias notas nao e alternativa), com fornecedor, endereco e codigo.
    alternativas = []
    por_loja: dict[str, Oferta] = {}
    for o in sorted(utilizaveis, key=chave):
        assinatura = normalizar(o.estabelecimento) + "|" + normalizar(o.endereco)
        if assinatura in por_loja:
            continue
        por_loja[assinatura] = o
        equivalente = None
        ppb = preco_por_base(o.preco, o.medida)
        if ppb is not None and medida_ref.total and medida_ref.base == o.medida.base:
            equivalente = round(ppb * medida_ref.total, 2)
        alternativas.append({
            "produto_portal": o.descricao,
            "embalagem": (f"{o.medida.total:g} {o.medida.base}" if o.medida.total else ""),
            "preco": o.preco,
            "preco_na_medida_da_planilha": equivalente,
            "economia_vs_atual": (round(preco_atual - equivalente, 2)
                                  if equivalente is not None and preco_atual is not None
                                  else None),
            "fornecedor": o.estabelecimento,
            "endereco": o.endereco,
            "municipio": o.municipio,
            "distancia_km": o.distancia_km,
            "data_venda": o.data_venda,
            "gtin": o.gtin,
            "confianca": getattr(o, "_conf", ""),
            "latitude": o.latitude,
            "longitude": o.longitude,
            "url_mapa": (f"https://www.google.com/maps/search/?api=1&query="
                         f"{o.latitude},{o.longitude}"
                         if o.latitude is not None and o.longitude is not None else ""),
        })
        if len(alternativas) >= max_alternativas:
            break
    linha["alternativas"] = alternativas
    linha["fornecedores_distintos"] = len(por_loja)

    ppb_ref = linha["preco_atual_por_base"]
    ppb_melhor = preco_por_base(melhor.preco, melhor.medida)
    if so_baixa:
        # Criterio: BAIXA nao entra no calculo sem sinalizacao. Aqui ele nem
        # entra: o preco fica visivel como referencia, a economia fica em branco.
        obs(
            "casamento fraco (BAIXA): nenhuma oferta bate marca e medida. Preco "
            "exibido apenas como referencia -- economia NAO calculada, conferir a mao"
        )
    elif (ppb_ref is not None and ppb_melhor is not None and medida_ref.total
          and medida_ref.base == melhor.medida.base):
        # bases iguais e obrigatorio: R$/ml nunca se compara com R$/g. Sem esta
        # guarda, 200 g de leite em po viravam "equivalente" a 1 L de leite.
        # economia projetada para a embalagem da planilha
        equivalente = ppb_melhor * medida_ref.total
        linha["economia_unitaria"] = round((preco_atual or 0) - equivalente, 4)
        linha["preco_equivalente_na_medida_da_planilha"] = round(equivalente, 4)
        if melhor.medida.ambigua or medida_ref.ambigua:
            obs(
                "descricao ambigua (unidade x fardo, ex. '500ML - 24X500ML'): lido "
                "como unidade, que e a leitura conservadora -- conferir a embalagem"
            )
    elif (preco_atual is not None and medida_ref.base and melhor.medida.base
          and medida_ref.base != melhor.medida.base):
        # Bases diferentes = categorias diferentes: 1 L de leite liquido nao se
        # compara com 400 g de leite EM PO. Nao ha economia a declarar.
        obs(
            "unidades incompativeis: a planilha pede %s e a melhor oferta e em %s "
            "(%s) -- provavelmente outro tipo de produto; economia NAO calculada"
            % (medida_ref.base, melhor.medida.base, melhor.descricao[:40])
        )
    elif preco_atual is not None:
        linha["economia_unitaria"] = round(preco_atual - melhor.preco, 4)
        if not comparaveis(medida_ref, melhor.medida):
            obs(
                "medidas nao comparaveis (volume/peso divergente ou ausente) -- "
                "economia calculada sobre preco absoluto, conferir manualmente"
            )
    if preco_atual and linha["economia_unitaria"] is not None:
        linha["economia_percentual"] = round(100 * linha["economia_unitaria"] / preco_atual, 2)

    if linha["distancia_km"] is None and linha["municipio_menor_preco"]:
        obs("distancia n/d")
    if uf.upper() == "AM" and linha["distancia_km"] and normalizar(melhor.municipio) != base_norm:
        obs("distancia em linha reta; no AM confirmar acesso (muitos municipios "
            "so por via fluvial)")
    linha["observacao"] = " | ".join(_obs)
    return linha


COLUNAS_SAIDA = [
    ("descricao_planilha", "Produto (planilha)", None),
    ("medida_planilha", "Medida", None),
    ("preco_atual", "Preco atual", "R$ #,##0.00"),
    ("menor_preco_municipio", "Menor no municipio", "R$ #,##0.00"),
    ("estabelecimento_municipio", "Estabelecimento (municipio)", None),
    ("menor_preco_estado", "Menor no estado", "R$ #,##0.00"),
    ("descricao_oferta", "Produto encontrado no portal", None),
    ("medida_oferta", "Embalagem encontrada", None),
    ("municipio_menor_preco", "Municipio do menor preco", None),
    ("estabelecimento_menor_preco", "Fornecedor (menor preco)", None),
    ("endereco_menor_preco", "Endereco do fornecedor", None),
    ("gtin", "Codigo de busca (GTIN)", None),
    ("fornecedores_distintos", "Fornecedores com o item", "0"),
    ("distancia_km", "Distancia (km, linha reta)", "#,##0.0"),
    ("preco_equivalente_na_medida_da_planilha", "Equivalente na medida da planilha", "R$ #,##0.00"),
    ("economia_unitaria", "Economia unitaria", "R$ #,##0.00"),
    ("economia_percentual", "Economia %", "0.00"),
    ("data_venda", "Data da venda (NFC-e)", None),
    ("confianca_match", "Confianca do match", None),
    ("ofertas_encontradas", "Ofertas", "0"),
    ("ofertas_descartadas_ruido", "Descartadas (outro produto)", "0"),
    ("ofertas_descartadas_outlier", "Descartadas (preco fora da curva)", "0"),
    ("termo_consultado", "Termo consultado", None),
    ("observacao", "Observacao", None),
]


def escrever_xlsx(resultados: list[dict], caminho: str, contexto: dict) -> None:
    """Planilha anotada: aba Comparativo + aba Resumo + aba Auditoria."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Comparativo"

    cab = Font(bold=True, color="FFFFFF")
    fundo = PatternFill("solid", fgColor="2E3A8C")
    for j, (_, rotulo, _) in enumerate(COLUNAS_SAIDA, 1):
        c = ws.cell(row=1, column=j, value=rotulo)
        c.font, c.fill = cab, fundo
        c.alignment = Alignment(vertical="center", wrap_text=True)

    verde = PatternFill("solid", fgColor="E1EFE7")
    amarelo = PatternFill("solid", fgColor="F5EBD8")
    cinza = PatternFill("solid", fgColor="EFEFEF")

    for i, r in enumerate(resultados, 2):
        for j, (chave, _, fmt) in enumerate(COLUNAS_SAIDA, 1):
            c = ws.cell(row=i, column=j, value=r.get(chave))
            if fmt and isinstance(r.get(chave), (int, float)):
                c.number_format = fmt
        conf = r.get("confianca_match")
        eco = r.get("economia_percentual") or 0
        idx_conf = [k for k, _, _ in COLUNAS_SAIDA].index("confianca_match") + 1
        idx_eco = [k for k, _, _ in COLUNAS_SAIDA].index("economia_percentual") + 1
        if conf == "BAIXA":
            ws.cell(row=i, column=idx_conf).fill = amarelo
        elif conf == "NAO_ENCONTRADO":
            ws.cell(row=i, column=idx_conf).fill = cinza
        if eco > 0:
            ws.cell(row=i, column=idx_eco).fill = verde

    larguras = {1: 42, 2: 12, 5: 30, 7: 40, 9: 24, 10: 30, 15: 22, 18: 22,
                19: 24, 20: 28, 21: 46}
    for j in range(1, len(COLUNAS_SAIDA) + 1):
        ws.column_dimensions[get_column_letter(j)].width = larguras.get(j, 16)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUNAS_SAIDA))}{max(2, len(resultados)+1)}"

    # ---- Resumo ----
    rs = wb.create_sheet("Resumo")
    achados = [r for r in resultados if r["confianca_match"] != "NAO_ENCONTRADO"]
    baixa = [r for r in resultados if r["confianca_match"] == "BAIXA"]
    positivos = [r for r in resultados if (r.get("economia_unitaria") or 0) > 0]
    fora = [r for r in positivos
            if normalizar(r.get("municipio_menor_preco") or "") != normalizar(contexto.get("municipio", ""))]
    municipios = sorted({r.get("municipio_menor_preco") for r in achados if r.get("municipio_menor_preco")})

    linhas_resumo = [
        ("UF consultada", contexto.get("uf", "")),
        ("Municipio de referencia", contexto.get("municipio", "")),
        ("Portal", contexto.get("portal", "")),
        ("Data da consulta", contexto.get("data", "")),
        ("", ""),
        ("Itens na planilha", len(resultados)),
        ("Itens com oferta encontrada", len(achados)),
        ("Itens sem oferta", len(resultados) - len(achados)),
        ("Itens com confianca BAIXA (conferir a mao)", len(baixa)),
        ("Itens com economia positiva", len(positivos)),
        ("Economia unitaria somada (R$)", round(sum(r["economia_unitaria"] for r in positivos), 2)),
        ("Itens cujo menor preco esta em OUTRO municipio", len(fora)),
        ("", ""),
        ("Municipios que apareceram", ", ".join(municipios) or "-"),
    ]
    for i, (k, v) in enumerate(linhas_resumo, 1):
        rs.cell(row=i, column=1, value=k).font = Font(bold=bool(k) and v != "")
        rs.cell(row=i, column=2, value=v)
    rs.column_dimensions["A"].width = 46
    rs.column_dimensions["B"].width = 60

    aviso = rs.cell(row=len(linhas_resumo) + 2, column=1,
                    value="Precos vem de NFC-e ja emitida: e preco praticado no passado recente, "
                          "nao oferta vigente. Distancias sao em linha reta. Itens BAIXA precisam "
                          "de conferencia humana.")
    aviso.alignment = Alignment(wrap_text=True, vertical="top")
    rs.merge_cells(start_row=aviso.row, start_column=1, end_row=aviso.row + 2, end_column=2)

    # ---- Alternativas ----
    # Uma linha por OPCAO, nao por item: com fornecedor, endereco, distancia e
    # codigo de busca, para dar para ligar na loja ou refazer a busca no portal.
    al = wb.create_sheet("Alternativas")
    cols_alt = [
        ("Produto (planilha)", "descricao_planilha", 40, None),
        ("Opcao", "_ordem", 7, "0"),
        ("Produto encontrado", "produto_portal", 40, None),
        ("Embalagem", "embalagem", 12, None),
        ("Preco", "preco", 12, 'R$ #,##0.00'),
        ("Equivale a (na sua medida)", "preco_na_medida_da_planilha", 22, 'R$ #,##0.00'),
        ("Economia por unidade", "economia_vs_atual", 20, 'R$ #,##0.00'),
        ("Fornecedor", "fornecedor", 34, None),
        ("Endereco", "endereco", 52, None),
        ("Municipio", "municipio", 18, None),
        ("Distancia (km)", "distancia_km", 14, '#,##0.0'),
        ("Data da venda", "data_venda", 22, None),
        ("Codigo de busca (GTIN)", "gtin", 20, None),
        ("Confianca", "confianca", 12, None),
        ("Mapa", "url_mapa", 14, None),
    ]
    for j, (rotulo, _c, larg, _f) in enumerate(cols_alt, 1):
        c = al.cell(row=1, column=j, value=rotulo)
        c.font, c.fill = cab, fundo
        c.alignment = Alignment(vertical="center", wrap_text=True)
        al.column_dimensions[get_column_letter(j)].width = larg
    al.row_dimensions[1].height = 30

    linha_al = 2
    for r in resultados:
        alts = r.get("alternativas") or []
        if not alts:
            al.cell(row=linha_al, column=1, value=r.get("descricao_planilha"))
            al.cell(row=linha_al, column=3,
                    value="sem oferta correspondente no portal").fill = cinza
            linha_al += 1
            continue
        for n, alt in enumerate(alts, 1):
            dados = dict(alt)
            dados["descricao_planilha"] = r["descricao_planilha"] if n == 1 else ""
            dados["_ordem"] = n
            for j, (_r, chave, _l, fmt) in enumerate(cols_alt, 1):
                c = al.cell(row=linha_al, column=j, value=dados.get(chave))
                if fmt and isinstance(dados.get(chave), (int, float)):
                    c.number_format = fmt
            if n == 1:                       # a melhor opcao de cada item
                for j in range(1, len(cols_alt) + 1):
                    al.cell(row=linha_al, column=j).fill = verde
            if (alt.get("confianca") or "") == "BAIXA":
                al.cell(row=linha_al, column=14).fill = amarelo
            cel_mapa = al.cell(row=linha_al, column=15)
            if alt.get("url_mapa"):
                cel_mapa.value = "abrir mapa"
                cel_mapa.hyperlink = alt["url_mapa"]
                cel_mapa.font = Font(color="0563C1", underline="single")
            else:
                cel_mapa.value = ""
            linha_al += 1
    al.freeze_panes = "C2"
    al.auto_filter.ref = f"A1:{get_column_letter(len(cols_alt))}{max(2, linha_al - 1)}"

    # ---- Auditoria ----
    au = wb.create_sheet("Auditoria")
    for j, rotulo in enumerate(["Produto", "Termo consultado", "Ofertas",
                                "Descartadas (outro produto)", "Confianca", "Observacao"], 1):
        c = au.cell(row=1, column=j, value=rotulo)
        c.font, c.fill = cab, fundo
    for i, r in enumerate(resultados, 2):
        au.cell(row=i, column=1, value=r.get("descricao_planilha"))
        au.cell(row=i, column=2, value=r.get("termo_consultado"))
        au.cell(row=i, column=3, value=r.get("ofertas_encontradas"))
        au.cell(row=i, column=4, value=r.get("ofertas_descartadas_ruido"))
        au.cell(row=i, column=5, value=r.get("confianca_match"))
        au.cell(row=i, column=6, value=r.get("observacao"))
    for j, w in {1: 42, 2: 30, 3: 10, 4: 26, 5: 18, 6: 60}.items():
        au.column_dimensions[get_column_letter(j)].width = w
    au.freeze_panes = "A2"

    wb.save(caminho)


def escrever_pdf(resultados: list[dict], caminho: str, contexto: dict) -> None:
    """PDF explicativo A4: metodo, itens com maior economia e ressalvas."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                    TableStyle, KeepTogether)

    AZUL = colors.HexColor("#2E3A8C")
    CINZA = colors.HexColor("#5C5F78")
    est = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=est["Title"], fontSize=20, leading=24,
                        textColor=AZUL, alignment=0, spaceAfter=2)
    sub = ParagraphStyle("sub", parent=est["Normal"], fontSize=10, textColor=CINZA,
                         leading=14, spaceAfter=14)
    h2 = ParagraphStyle("h2", parent=est["Heading2"], fontSize=12.5,
                        textColor=AZUL, spaceBefore=14, spaceAfter=4)
    corpo = ParagraphStyle("corpo", parent=est["Normal"], fontSize=9.5, leading=14)
    peq = ParagraphStyle("peq", parent=est["Normal"], fontSize=8.5, leading=12,
                         textColor=CINZA)

    doc = SimpleDocTemplate(
        caminho, pagesize=A4, title="Comparativo de precos NFC-e",
        author="comparador-preco-sefaz",
        leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm, bottomMargin=16 * mm,
    )
    hist = []
    hist.append(Paragraph("Comparativo de precos", h1))
    hist.append(Paragraph(
        f"{contexto.get('portal','')} &nbsp;|&nbsp; UF {contexto.get('uf','')} "
        f"&nbsp;|&nbsp; referencia: {contexto.get('municipio','')} "
        f"&nbsp;|&nbsp; consulta: {contexto.get('data','')}", sub))

    achados = [r for r in resultados if r["confianca_match"] != "NAO_ENCONTRADO"]
    baixa = [r for r in resultados if r["confianca_match"] == "BAIXA"]
    ausentes = [r for r in resultados if r["confianca_match"] == "NAO_ENCONTRADO"]
    positivos = sorted(
        [r for r in resultados if (r.get("economia_unitaria") or 0) > 0],
        key=lambda r: -(r.get("economia_unitaria") or 0),
    )
    total = round(sum(r["economia_unitaria"] for r in positivos), 2)

    painel = [[
        Paragraph(f"<b>{len(resultados)}</b><br/><font size=7>itens analisados</font>", corpo),
        Paragraph(f"<b>{len(achados)}</b><br/><font size=7>com oferta</font>", corpo),
        Paragraph(f"<b>{len(positivos)}</b><br/><font size=7>com economia</font>", corpo),
        Paragraph(f"<b>R$ {total:.2f}</b><br/><font size=7>economia unitaria somada</font>", corpo),
    ]]
    t = Table(painel, colWidths=[doc.width / 4.0] * 4)
    t.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#D6D8E2")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E4E6EE")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    hist.append(t)

    hist.append(Paragraph("De onde vem esse preco", h2))
    hist.append(Paragraph(
        "Cada preco abaixo saiu de uma NFC-e (nota fiscal de consumidor eletronica) "
        "efetivamente emitida por um estabelecimento do estado e enviada a SEFAZ. "
        "Ou seja: e preco que alguem <b>realmente pagou</b> no passado recente, nao "
        "uma oferta anunciada. O portal do Amazonas mostra por padrao as vendas das "
        "ultimas 48 horas (ajustavel de 1 a 7 dias); a familia Preco da Hora trabalha "
        "com janela de ate 3 dias. O estabelecimento nao tem obrigacao de manter o "
        "preco exibido.", corpo))

    hist.append(Paragraph("Como os produtos foram casados", h2))
    hist.append(Paragraph(
        "O casamento e feito <b>por descricao</b>: o nucleo do nome do produto (marca "
        "inclusive) e enviado ao portal, e cada resultado recebe um nivel de confianca. "
        "<b>ALTA</b> = marca e medida coincidem. <b>MEDIA</b> = uma das duas coincide. "
        "<b>BAIXA</b> = so parte do nome bateu; nesses casos o preco aparece como "
        "referencia, mas a <b>economia nao e calculada</b>. Resultados que nao tem nenhuma "
        "palavra em comum com o produto pedido sao <b>descartados</b> antes de qualquer "
        "conta: os portais casam por trecho de texto, e 'CEBOLA' chega a devolver 'BOLA DE "
        "ISOPOR'.", corpo))
    hist.append(Paragraph(
        "Toda comparacao de valor e feita por <b>preco por unidade base</b> (R$/litro ou "
        "R$/quilo): 1 L a R$ 1,50 e 500 ml a R$ 0,75 custam o mesmo, e a ferramenta nao "
        "reporta economia onde ela nao existe. Volumes em unidades diferentes (1 L de leite "
        "liquido contra 400 g de leite em po) nao sao comparados. A coluna <i>portal</i> "
        "mostra o que o estabelecimento realmente vendeu, porque o menor preco por litro "
        "pode ser um fardo -- nao a embalagem que voce compra.", corpo))

    if positivos:
        hist.append(Paragraph("Onde ha economia", h2))
        dados = [["Produto", "Atual", "Menor", "Municipio", "km", "Economia", "Conf."]]
        for r in positivos[:22]:
            # o produto da planilha e, embaixo, o que o portal realmente vendeu:
            # sem isso um fardo "24 X 500ML." passa por uma garrafa de 1 litro
            achado = str(r.get("descricao_oferta") or "")
            rotulo = str(r["descricao_planilha"])[:52]
            if achado:
                rotulo += f'<br/><font size=6 color="#5C5F78">portal: {achado[:46]}</font>'
            dados.append([
                Paragraph(rotulo, peq),
                f"{r['preco_atual']:.2f}" if r.get("preco_atual") is not None else "-",
                f"{r['menor_preco_estado']:.2f}" if r.get("menor_preco_estado") is not None else "-",
                Paragraph(str(r.get("municipio_menor_preco") or "-")[:18], peq),
                f"{r['distancia_km']:.0f}" if r.get("distancia_km") is not None else "n/d",
                f"{r['economia_unitaria']:.2f}",
                r["confianca_match"][:5],
            ])
        tab = Table(dados, colWidths=[doc.width * x for x in
                                      (0.34, 0.09, 0.09, 0.18, 0.07, 0.11, 0.12)],
                    repeatRows=1)
        tab.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), AZUL),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (1, 1), (5, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F5F9")]),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E4E6EE")),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        hist.append(tab)
        hist.append(Spacer(1, 4))
        hist.append(Paragraph(
            "A coluna km e a distancia em <b>linha reta</b> entre o municipio de "
            "referencia e o do estabelecimento. Economia e por unidade, ja convertida "
            "para a medida da sua planilha.", peq))

    # Item a item, SEMPRE -- inclusive quando nao ha nenhuma economia, que e
    # justamente quando um relatorio so com "onde ha economia" sairia vazio.
    hist.append(Paragraph("Item a item", h2))
    linhas_tab = [["Produto", "Atual", "Encontrado no portal", "Equiv.", "Dif.", "Conf."]]
    for r in sorted(resultados, key=lambda x: -(x.get("economia_unitaria") or -9e9)):
        achado = str(r.get("descricao_oferta") or "")
        emb = str(r.get("medida_oferta") or "")
        if achado:
            texto_achado = achado[:40] + (f" [{emb}]" if emb else "")
        else:
            texto_achado = "nao encontrado no portal"
        equiv = r.get("preco_equivalente_na_medida_da_planilha")
        eco = r.get("economia_unitaria")
        linhas_tab.append([
            Paragraph(str(r["descricao_planilha"])[:44], peq),
            f"{r['preco_atual']:.2f}" if r.get("preco_atual") is not None else "-",
            Paragraph(texto_achado, peq),
            f"{equiv:.2f}" if equiv is not None else "-",
            f"{eco:+.2f}" if eco is not None else "n/c",
            r["confianca_match"][:5],
        ])
    tab2 = Table(linhas_tab, colWidths=[doc.width * x for x in
                                        (0.28, 0.09, 0.36, 0.09, 0.09, 0.09)], repeatRows=1)
    tab2.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F5F9")]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E4E6EE")),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    hist.append(tab2)
    hist.append(Spacer(1, 4))
    hist.append(Paragraph(
        "<b>Equiv.</b> = quanto custaria a embalagem da sua planilha ao preco por litro/quilo "
        "encontrado. <b>Dif.</b> = quanto voce economizaria (+) ou pagaria a mais (-) por "
        "unidade; <i>n/c</i> = nao calculada, porque o casamento e fraco ou as unidades sao "
        "incompativeis.", peq))

    # ---- Alternativas por item: onde comprar, com fornecedor e codigo ----
    com_alt = [r for r in resultados if (r.get("alternativas") or [])]
    if com_alt:
        hist.append(Paragraph("Onde comprar — alternativas por item", h2))
        hist.append(Paragraph(
            "Até 5 fornecedores por produto, do menor para o maior preço na unidade "
            "base, um por estabelecimento. O <b>código</b> é o GTIN: com ele você "
            "refaz a busca exata no portal, sem depender da descrição.", peq))
        hist.append(Spacer(1, 6))
        for r in com_alt:
            emb_ref = r.get("medida_planilha") or "sem medida"
            cab_item = (f"<b>{r['descricao_planilha'][:60]}</b> "
                        f"<font size=7 color='#5C5F78'>— você paga "
                        f"R$ {r['preco_atual']:.2f}"
                        f"{'' if not r.get('preco_atual') else ''} · {emb_ref} · "
                        f"{r.get('fornecedores_distintos', 0)} fornecedor(es) com o item"
                        f"</font>")
            dados_alt = [["#", "Produto no portal / fornecedor", "Emb.", "Preço",
                          "Equiv.", "km", "Código"]]
            for n, a in enumerate(r["alternativas"], 1):
                local = a.get("fornecedor") or "-"
                if a.get("municipio"):
                    local += f" · {a['municipio']}"
                dados_alt.append([
                    str(n),
                    Paragraph(f"{a['produto_portal'][:44]}"
                              f"<br/><font size=6 color='#5C5F78'>{local[:56]}</font>", peq),
                    a.get("embalagem") or "-",
                    f"{a['preco']:.2f}",
                    (f"{a['preco_na_medida_da_planilha']:.2f}"
                     if a.get("preco_na_medida_da_planilha") is not None else "-"),
                    (f"{a['distancia_km']:.0f}" if a.get("distancia_km") is not None else "n/d"),
                    Paragraph(f"<font size=6>{a.get('gtin') or '-'}</font>", peq),
                ])
            t_alt = Table(dados_alt, colWidths=[doc.width * x for x in
                                                (0.04, 0.40, 0.09, 0.09, 0.09, 0.06, 0.23)],
                          repeatRows=1)
            t_alt.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E4E6EE")),
                ("TEXTCOLOR", (0, 0), (-1, 0), AZUL),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7.5),
                ("ALIGN", (2, 1), (5, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#E8F4EC")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E4E6EE")),
                ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]))
            hist.append(KeepTogether([Paragraph(cab_item, corpo), Spacer(1, 3),
                                      t_alt, Spacer(1, 10)]))

    ressalvas = [
        "O preco vem de nota ja emitida: pode ter mudado depois, e o produto pode estar "
        "sem estoque.",
        "A distancia e em linha reta, nao rota. No Amazonas, confirme o acesso: muitos "
        "municipios so tem ligacao fluvial, e 300 km em linha reta podem ser mais de um "
        "dia de viagem.",
        "Preco menor em municipio distante nao inclui frete nem eventual diferenca de "
        "ICMS. A ferramenta compara preco de etiqueta, nao custo total de aquisicao.",
        "Esta analise nao recomenda trocar de fornecedor ou de municipio: ela apresenta "
        "economia e distancia lado a lado para a sua decisao.",
    ]
    descartaram = [r for r in resultados if (r.get("ofertas_descartadas_ruido") or 0) > 0]
    if descartaram:
        total_ruido = sum(r["ofertas_descartadas_ruido"] for r in descartaram)
        ressalvas.insert(0, f"{total_ruido} resultado(s) do portal foram descartados por nao "
                            f"terem nenhuma palavra em comum com o produto pedido -- o portal "
                            f"casa por trecho de texto, e 'CEBOLA' chega a devolver 'BOLA DE "
                            f"ISOPOR'. Esses precos nao entraram em nenhuma conta.")
    outliers = [r for r in resultados if (r.get("ofertas_descartadas_outlier") or 0) > 0]
    if outliers:
        total_out = sum(r["ofertas_descartadas_outlier"] for r in outliers)
        ressalvas.insert(0, f"{total_out} oferta(s) foram descartadas por ter preco fora da "
                            f"distribuicao do proprio produto -- NFC-e de brinde, cortesia ou "
                            f"ajuste fiscal (uma lata de refrigerante a R$ 0,01, por exemplo) "
                            f"e preco real na nota, mas nao e preco de mercado.")
    if baixa:
        ressalvas.insert(0, f"{len(baixa)} item(ns) ficaram com confianca BAIXA: o preco aparece "
                            f"como referencia, mas a economia NAO foi calculada para eles. "
                            f"Conferir a mao: " +
                            ", ".join(str(r['descricao_planilha'])[:40] for r in baixa[:8]) +
                            ("..." if len(baixa) > 8 else ""))
    if ausentes:
        ressalvas.insert(0, f"{len(ausentes)} item(ns) sem nenhuma oferta no portal: " +
                            ", ".join(str(r['descricao_planilha'])[:40] for r in ausentes[:8]) +
                            ("..." if len(ausentes) > 8 else ""))

    bloco = [Paragraph("Ressalvas", h2)]
    for r in ressalvas:
        bloco.append(Paragraph(f"&bull;&nbsp; {r}", corpo))
        bloco.append(Spacer(1, 3))
    hist.append(KeepTogether(bloco))

    doc.build(hist)


# --------------------------------------------------------------------------
# 8. Selftest -- valida a logica pura, sem rede
# --------------------------------------------------------------------------

def selftest() -> int:
    falhas = []

    def check(nome, cond, detalhe=""):
        if cond:
            print(f"  ok   {nome}")
        else:
            print(f"  FALHA {nome} {detalhe}")
            falhas.append(nome)

    print("normalizacao e termo de busca")
    check("acento removido", normalizar("Detergente Líquido") == "DETERGENTE LIQUIDO")
    check("ruido de embalagem removido",
          "PCT" not in termo_de_busca("ARROZ TIO JOAO PCT 5KG"))
    check("medida fora do termo",
          termo_de_busca("DETERGENTE YPE NEUTRO 500ML") == "DETERGENTE YPE NEUTRO",
          f"-> {termo_de_busca('DETERGENTE YPE NEUTRO 500ML')!r}")

    print("extracao de medida")
    check("1L -> 1000ml", extrair_medida("DETERGENTE 1L").total == 1000.0)
    check("500ML -> 500ml", extrair_medida("DETERGENTE 500ML").total == 500.0)
    check("5KG -> 5000g", extrair_medida("ARROZ 5KG").total == 5000.0)
    check("1,5L -> 1500ml", extrair_medida("REFRIGERANTE 1,5L").total == 1500.0)
    check("12x500ml -> 6000ml", extrair_medida("CERVEJA 12X500ML").total == 6000.0)
    check("sem medida -> None", extrair_medida("CEBOLA").total is None)

    print("preco por unidade base -- o nucleo da comparacao honesta")
    p_1l = preco_por_base(1.50, extrair_medida("DETERGENTE 1L"))
    p_500 = preco_por_base(0.75, extrair_medida("DETERGENTE 500ML"))
    check("1L a 1,50 == 500ml a 0,75 (mesmo R$/L)", abs(p_1l - p_500) < 1e-12,
          f"{p_1l} vs {p_500}")
    p_500_barato = preco_por_base(0.60, extrair_medida("DETERGENTE 500ML"))
    check("500ml a 0,60 e realmente mais barato", p_500_barato < p_1l)

    print("parse de preco")
    check("R$ 1.234,56", parse_preco("R$ 1.234,56") == 1234.56)
    check("1,50", parse_preco("1,50") == 1.50)
    check("float direto", parse_preco(2.5) == 2.5)
    check("vazio -> None", parse_preco("") is None)

    print("haversine (verificavel por definicao)")
    d = haversine(0, 0, 0, 1)
    check("1 grau de longitude no equador ~111.19 km", abs(d - 111.19) < 0.1, f"-> {d:.3f}")
    check("distancia de um ponto a si mesmo e 0", haversine(-3.1, -60.0, -3.1, -60.0) == 0)
    check("simetrica",
          abs(haversine(-3.1, -60.0, -7.1, -34.8) - haversine(-7.1, -34.8, -3.1, -60.0)) < 1e-9)

    print("comparabilidade e confianca")
    check("1L vs 1000ML comparaveis",
          comparaveis(extrair_medida("X 1L"), extrair_medida("X 1000ML")))
    check("1L vs 500ML nao comparaveis",
          not comparaveis(extrair_medida("X 1L"), extrair_medida("X 500ML")))
    check("marca+medida iguais -> ALTA",
          nivel_confianca("DETERGENTE YPE NEUTRO 500ML", "DETERGENTE YPE NEUTRO 500 ML") == "ALTA")
    check("medida divergente -> nao ALTA",
          nivel_confianca("DETERGENTE YPE 500ML", "DETERGENTE YPE 1L") != "ALTA")
    check("produto sem palavra em comum -> RUIDO (nao e casamento fraco)",
          nivel_confianca("DETERGENTE YPE 500ML", "SABAO EM PO OMO 1KG") == "RUIDO")
    check("mesma marca, medida diferente -> MEDIA, nunca RUIDO",
          nivel_confianca("DETERGENTE YPE MACA 500ML", "DETERGENTE YPE MACA 2L") == "MEDIA",
          f"-> {nivel_confianca('DETERGENTE YPE MACA 500ML', 'DETERGENTE YPE MACA 2L')}")
    check("outra marca e outro tamanho -> RUIDO (nao e o seu produto)",
          nivel_confianca("DETERGENTE YPE MACA 500ML", "DETERGENTE GUAMA 2L") == "RUIDO")
    # o termo curto totalmente contido numa descricao de outra coisa
    check("CEBOLA x SALGADINHO ... CEBOLA E SALSA -> nao entra como preco",
          nivel_confianca("CEBOLA", "SALGADINHO AKIMILHO 30 GR CEBOLA E SALSA") == "RUIDO",
          f"-> {nivel_confianca('CEBOLA', 'SALGADINHO AKIMILHO 30 GR CEBOLA E SALSA')}")

    print("armadilhas vistas no portal real (regressao)")
    # "CEBOLA" casa por substring com "BOLA DE ISOPOR" no Busca Preco AM
    check("CEBOLA x BOLA DE ISOPOR -> RUIDO",
          nivel_confianca("CEBOLA", "BOLA DE ISOPOR 25MM") == "RUIDO")
    m_amb = extrair_medida("DETERGENTE GUAMA MACA 500ML - 24X500ML")
    check("'500ML - 24X500ML' lido como unidade de 500ml, nao fardo de 12L",
          m_amb.total == 500.0 and m_amb.ambigua, f"-> {m_amb}")
    check("multipack legitimo continua multipack (12X500ML = 6000ml)",
          extrair_medida("CERVEJA 12X500ML").total == 6000.0)
    # so ofertas de outro produto => nada e reportado
    r_ruido = consolidar("CEBOLA", 4.99,
                         [Oferta("BOLA DE ISOPOR 25MM", 5.99, "LOJA X", "", "Manaus",
                                 medida=extrair_medida("BOLA DE ISOPOR 25MM"))],
                         "Manaus", "AM", None, geocode=False)
    check("cebola com so isopor -> NAO_ENCONTRADO", r_ruido["confianca_match"] == "NAO_ENCONTRADO")
    check("cebola com so isopor -> sem menor preco", r_ruido["menor_preco_estado"] is None)
    check("cebola com so isopor -> sem economia", r_ruido["economia_unitaria"] is None)
    # ml nunca se compara com g
    check("R$/ml nao se compara com R$/g",
          consolidar("LEITE ITAMBE INTEGRAL 1L", 5.49,
                     [Oferta("LEITE ITAMBE INTEGRAL PO 200G", 2.99, "LOJA Y", "", "Manaus",
                             medida=extrair_medida("LEITE ITAMBE INTEGRAL PO 200G"))],
                     "Manaus", "AM", None, geocode=False
                     ).get("preco_equivalente_na_medida_da_planilha") is None)

    # so BAIXA => preco de referencia, mas economia em branco
    check("marca parcial + medida divergente -> BAIXA",
          nivel_confianca("ARROZ TIO JOAO INTEGRAL 5KG", "ARROZ TIO BRANCO 1KG") == "BAIXA",
          f"-> {nivel_confianca('ARROZ TIO JOAO INTEGRAL 5KG', 'ARROZ TIO BRANCO 1KG')}")
    r_baixa = consolidar("ARROZ TIO JOAO INTEGRAL 5KG", 28.90,
                         [Oferta("ARROZ TIO BRANCO 1KG", 5.99, "LOJA Y", "", "Manaus",
                                 medida=extrair_medida("ARROZ TIO BRANCO 1KG"))],
                         "Manaus", "AM", None, geocode=False)
    check("so BAIXA -> confianca BAIXA", r_baixa["confianca_match"] == "BAIXA")
    check("so BAIXA -> economia NAO calculada", r_baixa["economia_unitaria"] is None,
          f"-> {r_baixa['economia_unitaria']}")
    check("so BAIXA -> sinalizado na observacao", "BAIXA" in r_baixa["observacao"])

    print("deteccao de colunas")
    cols = detectar_colunas(["Produto", "Preço Atual", "EAN", "Unidade"])
    check("descricao", cols["descricao"] == "Produto", f"-> {cols}")
    check("preco", cols["preco"] == "Preço Atual", f"-> {cols}")
    check("gtin", cols["gtin"] == "EAN", f"-> {cols}")

    print("consolidacao -- caso do detergente do enunciado")
    ofertas = [
        Oferta("DETERGENTE YPE NEUTRO 500ML", 0.75, "MERCADO A", "", "Itacoatiara",
               medida=extrair_medida("DETERGENTE YPE NEUTRO 500ML"), distancia_km=176.0),
        Oferta("DETERGENTE YPE NEUTRO 1L", 1.50, "MERCADO B", "", "Manaus",
               medida=extrair_medida("DETERGENTE YPE NEUTRO 1L")),
    ]
    r = consolidar("DETERGENTE YPE NEUTRO 1L", 1.50, ofertas, "Manaus", "AM", None, geocode=False)
    check("economia zero: 500ml a 0,75 nao e mais barato que 1L a 1,50",
          abs((r["economia_unitaria"] or 0)) < 1e-9, f"-> {r['economia_unitaria']}")

    ofertas2 = [
        Oferta("DETERGENTE YPE NEUTRO 1L", 0.90, "MERCADO C", "", "Itacoatiara",
               medida=extrair_medida("DETERGENTE YPE NEUTRO 1L"), distancia_km=176.0),
    ]
    r2 = consolidar("DETERGENTE YPE NEUTRO 1L", 1.50, ofertas2, "Manaus", "AM", None, geocode=False)
    check("economia real de 0,60 detectada", abs((r2["economia_unitaria"] or 0) - 0.60) < 1e-9,
          f"-> {r2['economia_unitaria']}")
    check("municipio distante reportado", r2["municipio_menor_preco"] == "Itacoatiara")
    check("distancia reportada", r2["distancia_km"] == 176.0)
    check("ressalva de acesso fluvial no AM", "fluvial" in r2["observacao"])

    r3 = consolidar("CEBOLA", 4.99, [], "Manaus", "AM", None, geocode=False)
    check("sem oferta -> NAO_ENCONTRADO", r3["confianca_match"] == "NAO_ENCONTRADO")

    print("coleta da pagina de produto do Preco da Hora PB (Next.js)")
    coleta_pb = {
        "_fonte": "metadado, deve ser ignorado",
        "DETERGENTE YPE NEUTRO": {
            "produto": "DETERGENTE YPE LIQ 500ML NEUTRO",
            "gtin": "7896098900208",
            "lojas": [
                ["COMERCIAL SOARES & ARAUJO LTDA", 1.99, "07/09/2026", "MANGABEIRA", "JOAO PESSOA"],
                ["MERCADO AJUBA LTDA", 2.19, "04/09/2026", "PEDRO GONDIM", "JOAO PESSOA"],
                {"loja": "OURO BOM", "preco": 2.09, "data": "07/09/2026",
                 "bairro": "ALTO DA BOA VISTA", "cidade": "BAYEUX"},
            ],
        },
        "SABAO EM PO OMO": {"produto": "SABAO PO OMO 400G", "gtin": "", "lojas": []},
    }
    ad = AdapterColetado("PB", -7.1195, -34.8450, coleta_pb)
    of_pb = ad.buscar("DETERGENTE YPE NEUTRO")
    check("le as lojas da pagina de produto", len(of_pb) == 3, f"-> {len(of_pb)}")
    check("aceita loja como lista e como dicionario",
          sorted(o.preco for o in of_pb) == [1.99, 2.09, 2.19])
    check("GTIN do produto vai para todas as lojas",
          all(o.gtin == "7896098900208" for o in of_pb))
    check("medida sai do nome do produto", of_pb[0].medida.total == 500.0)
    check("fonte declara que o preco e media do lojista",
          "media do lojista" in of_pb[0].fonte, f"-> {of_pb[0].fonte}")
    check("bairro e cidade preenchidos",
          of_pb[0].endereco == "MANGABEIRA" and of_pb[0].municipio == "JOAO PESSOA")
    check("produto sem loja -> lista vazia", ad.buscar("SABAO EM PO OMO") == [])
    check("chave de metadado (_fonte) nao e confundida com termo",
          ad.buscar("termo inexistente") == [])

    print("codigo e coordenada escondidos nos gatilhos do card (AM)")
    card_falso = ('<div class="card"><a onclick="findByGtin(7893500018469);">x</a>'
                  '<a href="#modal-map" onclick="javascript:refreshMap( -59.9931023,'
                  ' -3.05139623399998);">mapa</a></div>')
    check("GTIN vem de findByGtin", AdapterAM._gtin(card_falso) == "7893500018469")
    lat, lon = AdapterAM._coordenadas(card_falso)
    check("coordenada e (lat, lon), na ordem trocada do portal",
          abs(lat + 3.05139623399998) < 1e-9 and abs(lon + 59.9931023) < 1e-9,
          f"-> {lat}, {lon}")
    # regressao: o fallback por digitos soltos capturava a coordenada como EAN-14
    check("card SEM findByGtin nao inventa codigo a partir da coordenada",
          AdapterAM._gtin('<div class="card"><a onclick="javascript:refreshMap('
                          ' -59.99, -3.05139623399998);">m</a></div>') == "",
          f"-> {AdapterAM._gtin(chr(60) + 'div><a onclick=refreshMap(-59.99,-3.05139623399998)></a></div>')!r}")

    print("outlier de NFC-e (lata de Coca a R$ 0,01 vista no portal do AM)")
    fora, med = descartar_outliers([6.0, 5.5, 6.5, 5.9, 0.01])
    check("R$ 0,01 entre precos de ~R$ 6 e outlier", fora == [4], f"-> {fora}, mediana {med}")
    check("nada descartado com amostra pequena", descartar_outliers([6.0, 0.01])[0] == [])
    check("preco baixo legitimo NAO e descartado (metade da mediana)",
          descartar_outliers([6.0, 5.5, 6.5, 5.9, 3.0])[0] == [])
    # o caso real: a lata de 350ml a R$ 0,01 nao pode virar o "menor preco"
    ofertas_coca = [
        Oferta("REFRIGERANTE COCA COLA 2L", 8.99, "MERC A", "", "Manaus",
               medida=extrair_medida("REFRIGERANTE COCA COLA 2L")),
        Oferta("REFRIGERANTE COCA COLA 1,5L", 7.29, "MERC B", "", "Manaus",
               medida=extrair_medida("REFRIGERANTE COCA COLA 1,5L")),
        Oferta("REFRIGERANTE COCA COLA 600ML", 3.49, "MERC C", "", "Manaus",
               medida=extrair_medida("REFRIGERANTE COCA COLA 600ML")),
        Oferta("REFRIGERANTE COCA COLA LT 350ML", 2.99, "MERC D", "", "Manaus",
               medida=extrair_medida("REFRIGERANTE COCA COLA LT 350ML")),
        Oferta("REFRIGERANTE COCA COLA LT 350ML", 0.01, "RESTAURANTE X", "", "Manaus",
               medida=extrair_medida("REFRIGERANTE COCA COLA LT 350ML")),
    ]
    r_out = consolidar("REFRIGERANTE COCA COLA 1,5L", 7.49, ofertas_coca,
                       "Manaus", "AM", None, geocode=False)
    check("lata a R$ 0,01 nao vira o menor preco", r_out["menor_preco_estado"] != 0.01,
          f"-> {r_out['menor_preco_estado']}")
    check("outlier contabilizado", r_out["ofertas_descartadas_outlier"] == 1,
          f"-> {r_out['ofertas_descartadas_outlier']}")
    check("descarte de outlier declarado na observacao",
          "fora da distribuicao" in r_out["observacao"], f"-> {r_out['observacao'][:70]}")
    check("sem o outlier, nao ha economia de 99%",
          (r_out["economia_percentual"] or 0) < 90, f"-> {r_out['economia_percentual']}")

    print()
    if falhas:
        print(f"{len(falhas)} falha(s): {', '.join(falhas)}")
        return 1
    print("todos os testes da logica pura passaram")
    return 0


# --------------------------------------------------------------------------
# 9. CLI
# --------------------------------------------------------------------------

def doctor() -> int:
    """
    Diz se esta maquina esta pronta para usar a skill, e o que fazer se nao.

    Existe porque quem instala isto em outro computador erra sempre nas mesmas
    duas coisas: dependencia faltando e Python velho. Melhor descobrir aqui, em
    dois segundos, do que no meio de uma consulta ao portal.
    """
    import importlib
    import platform

    problemas, avisos = [], []
    print("busca-preco -- verificacao de ambiente\n")

    print("Python")
    v = sys.version_info
    ok_py = v >= (3, 10)
    print("  %s %d.%d.%d (%s)" % ("ok  " if ok_py else "FALHA", v.major, v.minor,
                                  v.micro, platform.system()))
    if not ok_py:
        problemas.append("Python 3.10 ou mais novo (o codigo usa `X | Y` em anotacoes)")

    print("\nDependencias obrigatorias")
    for mod, pacote, para_que in (
        ("requests", "requests", "falar com os portais"),
        ("bs4", "beautifulsoup4", "ler o HTML do Amazonas"),
        ("lxml", "lxml", "acelerar a leitura do HTML"),
        ("openpyxl", "openpyxl", "ler e escrever planilhas"),
        ("reportlab", "reportlab", "gerar o PDF"),
    ):
        try:
            importlib.import_module(mod)
            print("  ok    %-16s %s" % (pacote, para_que))
        except ImportError:
            print("  FALTA %-16s %s" % (pacote, para_que))
            problemas.append("pip install " + pacote)

    print("\nOpcional")
    try:
        importlib.import_module("playwright")
        print("  ok    playwright       navegador proprio para a Paraiba")
    except ImportError:
        print("  -     playwright       so precisa se for usar --navegador na PB")
        avisos.append("playwright ausente: na PB, use a extensao do Chrome "
                      "(recomendado) ou instale com `pip install playwright` "
                      "+ `python -m playwright install chromium`")

    print("\nAcesso ao Busca Preco AM")
    try:
        import requests
        r = requests.get(AdapterAM.HOME, timeout=25,
                         headers={"User-Agent": UA})
        if r.status_code < 400:
            print("  ok    HTTP %s em %s" % (r.status_code, AdapterAM.HOME))
        else:
            print("  FALHA HTTP %s em %s" % (r.status_code, AdapterAM.HOME))
            problemas.append("o portal do AM respondeu HTTP %s" % r.status_code)
    except Exception as erro:
        print("  FALHA %s" % str(erro)[:90])
        problemas.append("sem acesso a buscapreco.sefaz.am.gov.br (rede? proxy?)")

    print("\nLogica pura")
    import io as _io
    saida = _io.StringIO()
    real, sys.stdout = sys.stdout, saida
    try:
        codigo = selftest()
    finally:
        sys.stdout = real
    linhas = saida.getvalue().splitlines()
    n_ok = sum(1 for l in linhas if l.strip().startswith("ok "))
    n_falha = sum(1 for l in linhas if "FALHA" in l)
    print("  %s %d verificacoes, %d falha(s)" % ("ok  " if codigo == 0 else "FALHA",
                                                 n_ok, n_falha))
    if codigo != 0:
        problemas.append("o autoteste falhou -- NAO use os resultados; rode "
                         "`--selftest` para ver quais")
        for l in linhas:
            if "FALHA" in l:
                print("     " + l.strip())

    print()
    for a in avisos:
        print("[aviso] " + a)
    if problemas:
        print("\nFALTA RESOLVER:")
        for p in problemas:
            print("  - " + p)
        print("\nAtalho para as dependencias:")
        print("  pip install requests beautifulsoup4 lxml openpyxl reportlab")
        return 1
    print("Tudo pronto. Amazonas funciona agora; para a Paraiba, veja o SKILL.md.")
    return 0


def construir_adapter(uf: str, municipio: str, raio: int, horas: int, paginas: int,
                      navegador: bool = False, espera_humano: int = 180,
                      perfil: str | None = None):
    uf = uf.upper()
    if uf == "AM":
        return AdapterAM(paginas=paginas), None
    if uf in ("PB", "BA"):
        coord = geocodificar(municipio, uf)
        if not coord:
            raise RuntimeError(
                f"Nao consegui geocodificar {municipio}/{uf}. O Preco da Hora exige "
                "latitude/longitude na consulta -- informe --lat e --lon."
            )
        if navegador:
            return (AdapterPrecoDaHoraNavegador(uf, coord[0], coord[1], raio, horas,
                                                paginas=paginas,
                                                espera_humano=espera_humano,
                                                perfil=perfil), coord)
        return AdapterPrecoDaHora(uf, coord[0], coord[1], raio, horas, paginas=paginas), coord
    raise SystemExit(
        f"UF {uf} nao suportada. Este comparador cobre AM e PB (e BA, mesmo "
        "codigo-base do PB). PR e AL tem API propria e mais facil; as demais UFs "
        "so tem o app Menor Preco Brasil, sem web e sem API publica."
    )


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Comparador de preco SEFAZ (NFC-e)")
    ap.add_argument("--uf", help="AM, PB ou BA")
    ap.add_argument("--municipio", default="", help="municipio de referencia")
    ap.add_argument("--planilha", help="caminho .xlsx/.csv")
    ap.add_argument("--saida", default="comparativo_precos", help="prefixo dos arquivos de saida")
    ap.add_argument("--raio", type=int, default=30, help="km (Preco da Hora, max 30)")
    ap.add_argument("--horas", type=int, default=72, help="janela de dados (Preco da Hora)")
    ap.add_argument("--paginas", type=int, default=2)
    ap.add_argument("--limite", type=int, default=0, help="processa so os N primeiros itens")
    ap.add_argument("--brutos", metavar="DIR",
                    help="grava a resposta crua de cada consulta neste diretorio, "
                         "para poder mostrar de onde veio cada numero")
    ap.add_argument("--listar-termos", action="store_true",
                    help="le a planilha e imprime, em JSON, os termos a consultar; "
                         "nao toca na rede. Use para coletar por fora (extensao do Chrome)")
    ap.add_argument("--ofertas-json", metavar="ARQ",
                    help="processa respostas cruas do portal ja coletadas, em vez de "
                         "consultar: {termo: [resposta de /produtos/, ...]}")
    ap.add_argument("--navegador", action="store_true",
                    help="PB: usa Chrome real; uma PESSOA conclui a verificacao "
                         "do Cloudflare na janela aberta")
    ap.add_argument("--espera-humano", type=int, default=180, metavar="SEG",
                    help="segundos de espera pela verificacao humana (padrao 180)")
    ap.add_argument("--perfil", metavar="DIR",
                    help="PB: pasta de perfil do Chrome para guardar a sessao; o "
                         "desafio do Cloudflare, resolvido uma vez, vale nas proximas")
    ap.add_argument("--doctor", action="store_true",
                    help="verifica se esta maquina esta pronta (dependencias, "
                         "Python, acesso ao portal) e diz o que falta")
    ap.add_argument("--selftest", action="store_true", help="testa a logica pura, sem rede")
    ap.add_argument("--smoke", action="store_true", help="1 consulta real para validar acesso")
    ap.add_argument("--termo", default="detergente", help="termo do --smoke")
    args = ap.parse_args(argv)

    if args.doctor:
        return doctor()

    if args.selftest:
        return selftest()

    if not args.uf:
        ap.error("--uf e obrigatorio")

    padrao = {"AM": "Manaus", "PB": "Joao Pessoa", "BA": "Salvador"}
    municipio = args.municipio or padrao.get(args.uf.upper(), "")
    if not args.municipio:
        print(f"[aviso] --municipio nao informado; usando {municipio}", file=sys.stderr)

    if args.smoke:
        adapter, _ = construir_adapter(args.uf, municipio, args.raio, args.horas, 1,
                                   args.navegador, args.espera_humano, args.perfil)
        try:
            ofertas = adapter.buscar(termo_de_busca(args.termo))
        except RuntimeError as erro:
            # portal fora do ar, rota mudada ou anti-bot: condicao esperada de um
            # teste de acesso, nao merece traceback
            print(f"acesso a {args.uf.upper()} FALHOU: {erro}", file=sys.stderr)
            return 2
        finally:
            if hasattr(adapter, "fechar"):
                adapter.fechar()
        print(f"{len(ofertas)} oferta(s) para {args.termo!r} em {args.uf.upper()}")
        for of in ofertas[:10]:
            print(f"  R$ {of.preco:>8.2f}  {of.descricao[:48]:<48}  "
                  f"{of.estabelecimento[:24]:<24} {of.municipio}")
        return 0 if ofertas else 2

    if not args.planilha:
        ap.error("--planilha e obrigatorio (ou use --smoke / --selftest)")

    linhas, cabecalhos = ler_planilha(args.planilha)
    cols = detectar_colunas(cabecalhos)
    if not cols["descricao"] or not cols["preco"]:
        print(f"Nao identifiquei as colunas. Cabecalhos: {cabecalhos}", file=sys.stderr)
        print(f"Detectado: {cols}", file=sys.stderr)
        return 1
    print(f"{len(linhas)} linha(s); descricao={cols['descricao']!r} preco={cols['preco']!r}",
          file=sys.stderr if args.listar_termos else sys.stdout)

    alvo_termos = linhas[: args.limite] if args.limite else linhas
    if args.listar_termos:
        # Contrato com quem coleta por fora (extensao do Chrome, na PB): a lista
        # do que consultar, ja normalizada, sem tocar na rede.
        saida_termos = []
        vistos_t = set()
        for linha in alvo_termos:
            desc = str(linha.get(cols["descricao"]) or "").strip()
            if not desc:
                continue
            termo = termo_de_busca(desc)
            saida_termos.append({
                "descricao_planilha": desc,
                "preco_atual": parse_preco(linha.get(cols["preco"])),
                "termo": termo,
                "ja_pedido": termo in vistos_t,   # repetido: consulte uma vez so
            })
            vistos_t.add(termo)
        print(json.dumps({
            "uf": args.uf.upper(),
            "municipio": municipio,
            "raio_km": min(args.raio, 30),
            "horas": args.horas,
            "itens": saida_termos,
            "termos_unicos": sorted(vistos_t),
        }, ensure_ascii=False, indent=2))
        return 0

    if args.ofertas_json:
        with open(args.ofertas_json, encoding="utf-8") as fh:
            coletado = json.load(fh)
        coord_base = geocodificar(municipio, args.uf)
        if not coord_base:
            print(f"[aviso] sem geocodificacao de {municipio}/{args.uf}: distancias sairao n/d",
                  file=sys.stderr)
        adapter = AdapterColetado(args.uf, (coord_base or (0.0, 0.0))[0],
                                  (coord_base or (0.0, 0.0))[1], coletado)
        print(f"processando coleta externa: {len(coletado)} termo(s) em {args.ofertas_json}")
    else:
        adapter, coord_base = construir_adapter(args.uf, municipio, args.raio, args.horas,
                                               args.paginas, args.navegador,
                                               args.espera_humano, args.perfil)
        if coord_base is None:
            coord_base = geocodificar(municipio, args.uf)

    alvo = linhas[: args.limite] if args.limite else linhas
    cache: dict[str, list[Oferta]] = {}
    resultados = []
    abortou = False
    for i, linha in enumerate(alvo, 1):
        desc = str(linha.get(cols["descricao"]) or "").strip()
        if not desc:
            continue
        preco = parse_preco(linha.get(cols["preco"]))
        termo = termo_de_busca(desc)
        print(f"[{i}/{len(alvo)}] {desc[:56]}  ->  {termo!r}")
        if termo in cache:
            ofertas = cache[termo]
        elif abortou:
            ofertas = []                 # a sessao ja falhou: nao insiste no portal
        else:
            try:
                ofertas = adapter.buscar(termo)
            except Exception as erro:
                print(f"    erro: {erro}", file=sys.stderr)
                ofertas = []
                # Falha de SESSAO (anti-bot, rota morta, portal fora do ar) nao se
                # resolve no item seguinte. Repetir 14 vezes uma espera de 8
                # minutos seria martelar a SEFAZ por duas horas. Aborta a coleta e
                # ainda assim gera o relatorio, marcando o que ficou sem consulta.
                if not getattr(adapter, "csrf", "ok") and not abortou:
                    abortou = True
                    print("    coleta interrompida: sessao com o portal nao foi "
                          "estabelecida; os itens restantes ficam sem consulta",
                          file=sys.stderr)
            cache[termo] = ofertas
        resultados.append(
            consolidar(desc, preco, ofertas, municipio, args.uf, coord_base)
        )

    contexto = {
        "uf": args.uf.upper(),
        "municipio": municipio,
        "portal": {"AM": "Busca Preco AM", "PB": "Preco da Hora PB",
                   "BA": "Preco da Hora BA"}[args.uf.upper()],
        "data": time.strftime("%d/%m/%Y %H:%M"),
    }
    base_saida = os.path.splitext(args.saida)[0]
    with open(base_saida + ".json", "w", encoding="utf-8") as fh:
        json.dump(resultados, fh, ensure_ascii=False, indent=2, default=str)

    if args.brutos:
        # A resposta crua de cada consulta, para que qualquer numero do relatorio
        # possa ser rastreado ate o que o portal respondeu.
        os.makedirs(args.brutos, exist_ok=True)
        ext = "html" if args.uf.upper() == "AM" else "json"
        for n, texto in enumerate(getattr(adapter, "_raw", []), 1):
            with open(os.path.join(args.brutos, f"consulta_{n:03d}.{ext}"),
                      "w", encoding="utf-8") as fh:
                fh.write(texto)
        print(f"respostas brutas ({len(getattr(adapter, '_raw', []))}) em {args.brutos}/")
    try:
        escrever_xlsx(resultados, base_saida + ".xlsx", contexto)
        print(f"planilha em {base_saida}.xlsx")
    except Exception as erro:
        print(f"[aviso] nao consegui gravar o xlsx: {erro}", file=sys.stderr)
    try:
        escrever_pdf(resultados, base_saida + ".pdf", contexto)
        print(f"pdf em {base_saida}.pdf")
    except Exception as erro:
        print(f"[aviso] nao consegui gravar o pdf: {erro}", file=sys.stderr)
    if abortou:
        print("[ATENCAO] a coleta foi interrompida: o relatorio abaixo NAO cobre "
              "todos os itens da planilha", file=sys.stderr)
    achados = sum(1 for r in resultados if r["confianca_match"] != "NAO_ENCONTRADO")
    economia = sum(r["economia_unitaria"] or 0 for r in resultados
                   if (r["economia_unitaria"] or 0) > 0)
    print(f"\n{achados}/{len(resultados)} itens com oferta; "
          f"economia unitaria somada R$ {economia:.2f}")
    print(f"dados brutos em {base_saida}.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
