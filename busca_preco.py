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

# Assets de marca em marca/: logo em vetor gerada do SVG oficial. Sem eles o
# relatorio ainda sai, na fonte de reserva e sem logo, e avisa no rodape.
try:
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "marca"))
    import logo_aions as _LOGO
except Exception:
    _LOGO = None
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
    r"(?<![\d,\.])(\d{1,5}(?:[.,]\d{1,3})?)\s*"
    r"(ml|l|lt|litros?|kg|k|g|gr|gramas?|grama|mg|quilo)(?![a-z])",
    re.IGNORECASE,
)
# "12x500ml", "6 x 1L" -> multipack
_RE_MULTI = re.compile(r"(\d{1,3})\s*[x\*]\s*(\d{1,5}(?:[.,]\d{1,3})?)\s*(ml|l|lt|kg|g|gr)\b", re.I)

# Virgula decimal PERDIDA na digitacao: "SUCO DEL VALLE UVA 1 5L" e 1,5 L, nao
# 1 unidade de 5 L. Visto no Busca Preco AM, onde a descricao e digitada a mao.
# Lido como 5 L, o preco por litro despenca e inventa economia.
_RE_DECIMAL_COM_ESPACO = re.compile(
    r"(?<![\d,\.])(\d{1,2})\s+(\d{1,2})\s*(ml|l|lt|kg|g|gr)(?![a-z])", re.IGNORECASE)

# Contagem de unidades: "06 UN", "12 UNIDADES", "DUZIA", "MEIA DUZIA".
# Ovos, sabonetes e fardos sao vendidos assim, e comparar 12 ovos com 6 ovos
# pelo preco de etiqueta gera "67% de economia" que nao existe.
_RE_CONTAGEM = re.compile(
    r"(?<![\d,\.])(\d{1,3})\s*(?:un|und|unid|unidades?|uni)(?![a-z])", re.IGNORECASE)
_RE_DUZIA = re.compile(r"\b(meia\s+d[uú]zia|d[uú]zia)\b", re.IGNORECASE)


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

    # Virgula decimal perdida ANTES de tudo: "1 5L" tem de virar "1,5L", senao o
    # _RE_MEDIDA le "5L" e o volume sai 3x maior do que e.
    m_dec = _RE_DECIMAL_COM_ESPACO.search(texto)
    if m_dec and not _RE_MULTI.search(texto):
        valor = float("%s.%s" % (m_dec.group(1), m_dec.group(2)))
        base, fator = UNIDADES[m_dec.group(3).lower()]
        return Medida(valor * fator, base, 1, ambigua=True)

    # Contagem de unidades, quando nao ha volume nem peso na descricao
    if not _RE_MEDIDA.search(texto) and not _RE_MULTI.search(texto):
        m_duz = _RE_DUZIA.search(texto)
        if m_duz:
            return Medida(6.0 if "MEIA" in m_duz.group(1).upper() else 12.0, "un", 1)
        m_cont = _RE_CONTAGEM.search(texto)
        if m_cont:
            return Medida(float(m_cont.group(1)), "un", 1)

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
            lat = lon = dist = None
            if isinstance(linha, dict):
                loja, preco = linha.get("loja"), linha.get("preco")
                data, bairro = linha.get("data"), linha.get("bairro")
                cidade = linha.get("cidade")
                # a PB reescrita entrega coordenada e distancia por loja
                lat = parse_preco(linha.get("latitude"))
                lon = parse_preco(linha.get("longitude"))
                dist = parse_preco(linha.get("distancia"))
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
            # a distancia sai da coordenada da loja, nao do centroide do
            # municipio; a do portal fica como alternativa
            if dist is None and lat is not None and lon is not None:
                dist = round(haversine(self.lat, self.lon, lat, lon), 1)
            ofertas.append(Oferta(
                descricao=nome, preco=valor, gtin=gtin,
                estabelecimento=str(loja),
                endereco=str(bairro or ""),
                municipio=str(cidade or ""),
                data_venda=str(data or "")[:10],
                latitude=lat, longitude=lon, distancia_km=dist,
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
    achado = {"descricao": None, "preco": None, "gtin": None, "unidade": None,
              "fornecedor": None}
    for cab in cabecalhos:
        n = normalizar(cab).lower()
        # Fornecedor e testado ANTES e sai do laco: "Fornecedor Atual" contem
        # "atual", que tambem casa com "preco atual" de CAB_PRECO.
        if achado["fornecedor"] is None and any(
                k in n for k in ("fornecedor", "loja", "estabelecimento", "vendedor")):
            achado["fornecedor"] = cab
            continue
        if achado["descricao"] is None and any(k in n for k in CAB_DESC):
            achado["descricao"] = cab
        if achado["preco"] is None and any(k in n for k in CAB_PRECO):
            achado["preco"] = cab
        if achado["gtin"] is None and any(k in n for k in ("gtin", "ean", "barras", "codigo de barras")):
            achado["gtin"] = cab
        if achado["unidade"] is None and any(k in n for k in ("unidade", "embalagem", "volume", "un")):
            achado["unidade"] = cab
    return achado


def mesmo_fornecedor(a: str, b: str) -> bool:
    """
    O fornecedor da planilha e o do portal sao o mesmo?

    A razao social do portal quase nunca bate com o nome usado na planilha:
    "Higiluz Comercial" contra "HIGILUZ COMERCIO DE PRODUTOS LTDA". Compara por
    palavras significativas, descartando as genericas -- senao "COMERCIO" ou
    "LTDA" casariam meio Manaus. Exige duas palavras em comum, ou uma so quando
    ela e distintiva (5 letras ou mais).
    """
    GENERICAS = {"COMERCIAL", "COMERCIO", "LTDA", "ME", "EPP", "EIRELI", "SA",
                 "DE", "DA", "DO", "E", "MERCADO", "MERCADINHO", "SUPERMERCADO",
                 "DISTRIBUIDORA", "ATACADO", "ATACADAO", "PRODUTOS", "ALIMENTOS",
                 "LOJA", "VAREJAO", "CASA", "EMPORIO"}
    ta = {p for p in normalizar(a).split() if p not in GENERICAS and len(p) > 1}
    tb = {p for p in normalizar(b).split() if p not in GENERICAS and len(p) > 1}
    if not ta or not tb:
        return False
    comuns = ta & tb
    if len(comuns) >= 2:
        return True
    return any(len(p) >= 5 for p in comuns)


# --------------------------------------------------------------------------
# 7. Consolidacao
# --------------------------------------------------------------------------

def consolidar(descricao: str, preco_atual: float | None, ofertas: list[Oferta],
               municipio_base: str, uf: str, coord_base: tuple[float, float] | None,
               geocode: bool = True, max_alternativas: int = 5,
               fornecedor_atual: str = "") -> dict:
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
        "medida_planilha": (("%g %s" % (medida_ref.total, medida_ref.base))
                            if medida_ref.total else ""),
        "preco_atual_por_base": preco_por_base(preco_atual, medida_ref),
        "ofertas_encontradas": len(ofertas),
        "confianca_match": "NAO_ENCONTRADO",
        "menor_preco_municipio": None,
        "estabelecimento_municipio": "",
        "menor_preco_estado": None,
        "descricao_oferta": "",
        "gtin": "",
        "alternativas": [],
        "fornecedor_atual": fornecedor_atual,
        "preco_fornecedor_atual_no_portal": None,
        "variacao_no_fornecedor_atual": None,
        "grupo": "SEM_PRECO",
        "alt1_preco": None, "alt1_equivalente": None,
        "alt1_fornecedor": "", "alt1_municipio": "", "alt1_km": None,
        "alt2_preco": None, "alt2_equivalente": None,
        "alt2_fornecedor": "", "alt2_municipio": "", "alt2_km": None,
        "alt3_preco": None, "alt3_equivalente": None,
        "alt3_fornecedor": "", "alt3_municipio": "", "alt3_km": None,
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
            "sem nenhuma palavra em comum (ex.: %s). O portal casa por aproximação "
            "de texto: 'CEBOLA' puxa 'COLA' e 'BOLA', e tudo foi descartado"
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
            "%d oferta(s) descartada(s) por preço fora da distribuição do produto "
            "(ex.: %s a R$ %.2f, contra a mediana das outras). NFC-e de brinde ou "
            "ajuste fiscal não é preço de mercado"
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

    # As alternativas tambem entram ACHATADAS na linha, para caber na aba
    # Comparativo. A posicao 0 da lista e o proprio "menor no estado", que ja
    # tem colunas proprias -- entao "Alternativa 1" e a SEGUNDA melhor, e o
    # rotulo da coluna diz isso. Sem esse cuidado alguem leria a mesma loja
    # duas vezes e acharia que sao duas fontes para o mesmo preco.
    for n in (1, 2, 3):
        alt = alternativas[n] if len(alternativas) > n else None
        linha["alt%d_preco" % n] = alt["preco"] if alt else None
        linha["alt%d_equivalente" % n] = alt["preco_na_medida_da_planilha"] if alt else None
        linha["alt%d_fornecedor" % n] = alt["fornecedor"] if alt else ""
        linha["alt%d_municipio" % n] = alt["municipio"] if alt else ""
        linha["alt%d_km" % n] = alt["distancia_km"] if alt else None

    ppb_ref = linha["preco_atual_por_base"]
    ppb_melhor = preco_por_base(melhor.preco, melhor.medida)
    if so_baixa:
        # Criterio: BAIXA nao entra no calculo sem sinalizacao. Aqui ele nem
        # entra: o preco fica visivel como referencia, a economia fica em branco.
        obs(
            "casamento fraco: nenhuma oferta bate marca e medida. O preço aparece "
            "como referência e a economia não foi calculada, confira à mão"
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
                "descrição ambígua (unidade ou fardo, ex. '500ML - 24X500ML'): lida "
                "como unidade, que é a leitura conservadora, confira a embalagem"
            )
    elif (preco_atual is not None and medida_ref.base and melhor.medida.base
          and medida_ref.base != melhor.medida.base):
        # Bases diferentes = categorias diferentes: 1 L de leite liquido nao se
        # compara com 400 g de leite EM PO. Nao ha economia a declarar.
        obs(
            "unidades incompatíveis: a planilha pede %s e a melhor oferta é em %s "
            "(%s), provavelmente outro tipo de produto, então a economia não foi "
            "calculada"
            % (medida_ref.base, melhor.medida.base, melhor.descricao[:40])
        )
    elif preco_atual is not None and medida_ref.total and not melhor.medida.total:
        # A planilha diz o tamanho, a oferta do portal nao. Subtrair os precos
        # de etiqueta compara coisas de tamanho desconhecido: foi assim que uma
        # linguica de R$ 8,75 sem gramatura virou "64% de economia" sobre um
        # pacote de 500 g. Sem a medida do outro lado nao ha economia a declarar.
        obs(
            "a oferta do portal não informa o tamanho (%s): sem isso não há como "
            "comparar com %s da sua planilha, então a economia não foi calculada"
            % (melhor.descricao[:38], linha["medida_planilha"])
        )
    elif preco_atual is not None:
        linha["economia_unitaria"] = round(preco_atual - melhor.preco, 4)
        if not comparaveis(medida_ref, melhor.medida):
            obs(
                "medidas não comparáveis (volume ou peso divergente), então a economia "
                "saiu do preço de etiqueta, confira manualmente"
            )
    if preco_atual and linha["economia_unitaria"] is not None:
        linha["economia_percentual"] = round(100 * linha["economia_unitaria"] / preco_atual, 2)

    if linha["distancia_km"] is None and linha["municipio_menor_preco"]:
        obs("distancia n/d")
    if uf.upper() == "AM" and linha["distancia_km"] and normalizar(melhor.municipio) != base_norm:
        obs("distancia em linha reta; no AM confirmar acesso (muitos municipios "
            "so por via fluvial)")
    # ------------------------------------------------------------------
    # O SEU fornecedor atual aparece no portal? Se aparece, da para ver o que
    # ele esta cobrando hoje contra o que a planilha registra -- e o unico jeito
    # honesto de dizer "aumentou no proprio fornecedor" sem inventar historico.
    # ------------------------------------------------------------------
    if fornecedor_atual:
        do_atual = [o for o in utilizaveis if mesmo_fornecedor(fornecedor_atual, o.estabelecimento)]
        if do_atual:
            no_portal = min(do_atual, key=chave)
            eq_atual = None
            ppb_at = preco_por_base(no_portal.preco, no_portal.medida)
            if ppb_at is not None and medida_ref.total and medida_ref.base == no_portal.medida.base:
                eq_atual = round(ppb_at * medida_ref.total, 2)
            referencia = eq_atual if eq_atual is not None else no_portal.preco
            linha["preco_fornecedor_atual_no_portal"] = referencia
            if preco_atual:
                linha["variacao_no_fornecedor_atual"] = round(referencia - preco_atual, 2)

    # ------------------------------------------------------------------
    # Grupo de decisao -- e o que organiza o relatorio
    # ------------------------------------------------------------------
    eco = linha["economia_unitaria"]
    var = linha["variacao_no_fornecedor_atual"]
    if linha["menor_preco_estado"] is None:
        linha["grupo"] = "SEM_PRECO"
    elif eco is None:
        linha["grupo"] = "CONFERIR"
    elif var is not None and var > 0.009 and eco <= 0.009:
        # o mercado nao esta mais barato, MAS o proprio fornecedor atual ja
        # cobra mais do que a planilha registra: alta de preco, nao oportunidade
        linha["grupo"] = "SUBIU_NO_ATUAL"
    elif eco > 0.009:
        linha["grupo"] = "TROCAR"
    else:
        linha["grupo"] = "MANTER"

    linha["observacao"] = " | ".join(_obs)
    return linha


COLUNAS_SAIDA = [
    # 1. o que voce tem hoje
    ("grupo", "Decisao", None),
    ("descricao_planilha", "Produto (planilha)", None),
    ("medida_planilha", "Embalagem", None),
    ("preco_atual", "Voce paga", "R$ #,##0.00"),
    ("fornecedor_atual", "Seu fornecedor", None),
    # 2. o mesmo fornecedor, no portal
    ("preco_fornecedor_atual_no_portal", "Seu fornecedor cobra hoje", "R$ #,##0.00"),
    ("variacao_no_fornecedor_atual", "Alta no seu fornecedor", "R$ #,##0.00"),
    # 3. a melhor opcao do mercado -- preco SEMPRE com municipio ao lado
    ("preco_equivalente_na_medida_da_planilha", "Melhor preco (na sua embalagem)", "R$ #,##0.00"),
    ("economia_unitaria", "Economia por unidade", "R$ #,##0.00"),
    ("economia_percentual", "Economia %", "0.0"),
    ("estabelecimento_menor_preco", "Fornecedor", None),
    ("municipio_menor_preco", "Municipio", None),
    ("distancia_km", "Distancia (km)", "#,##0.0"),
    ("endereco_menor_preco", "Endereco", None),
    ("menor_preco_estado", "Preco de etiqueta", "R$ #,##0.00"),
    ("descricao_oferta", "Produto encontrado no portal", None),
    ("medida_oferta", "Embalagem encontrada", None),
    ("data_venda", "Data da venda (NFC-e)", None),
    ("gtin", "Codigo de busca (GTIN)", None),
    # 4. alternativas -- cada uma com municipio proprio
    ("alt1_equivalente", "Alternativa 1: preco", "R$ #,##0.00"),
    ("alt1_fornecedor", "Alternativa 1: fornecedor", None),
    ("alt1_municipio", "Alternativa 1: municipio", None),
    ("alt1_km", "Alternativa 1: km", "#,##0.0"),
    ("alt2_equivalente", "Alternativa 2: preco", "R$ #,##0.00"),
    ("alt2_fornecedor", "Alternativa 2: fornecedor", None),
    ("alt2_municipio", "Alternativa 2: municipio", None),
    ("alt2_km", "Alternativa 2: km", "#,##0.0"),
    ("alt3_equivalente", "Alternativa 3: preco", "R$ #,##0.00"),
    ("alt3_fornecedor", "Alternativa 3: fornecedor", None),
    ("alt3_municipio", "Alternativa 3: municipio", None),
    ("alt3_km", "Alternativa 3: km", "#,##0.0"),
    # 5. rastreabilidade
    ("confianca_match", "Confianca", None),
    ("fornecedores_distintos", "Fornecedores com o item", "0"),
    ("ofertas_encontradas", "Ofertas lidas", "0"),
    ("ofertas_descartadas_ruido", "Descartadas: outro produto", "0"),
    ("ofertas_descartadas_outlier", "Descartadas: fora da curva", "0"),
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

    ROTULO_GRUPO = {
        "TROCAR": "1 TROCAR", "MANTER": "2 MANTER", "SUBIU_NO_ATUAL": "3 SUBIU",
        "CONFERIR": "4 CONFERIR", "SEM_PRECO": "4 SEM PRECO",
    }
    ORDEM_GRUPO = {"TROCAR": 0, "SUBIU_NO_ATUAL": 1, "MANTER": 2,
                   "CONFERIR": 3, "SEM_PRECO": 4}
    ordenados = sorted(resultados,
                       key=lambda r: (ORDEM_GRUPO.get(r.get("grupo"), 9),
                                      -(r.get("economia_unitaria") or 0)))
    for i, r in enumerate(ordenados, 2):
        for j, (chave, _, fmt) in enumerate(COLUNAS_SAIDA, 1):
            valor = r.get(chave)
            if chave == "grupo":
                valor = ROTULO_GRUPO.get(valor, valor)
            c = ws.cell(row=i, column=j, value=valor)
            if fmt and isinstance(r.get(chave), (int, float)):
                c.number_format = fmt
        # cor na coluna de decisao: verde troca, cinza mantem, vermelho subiu
        cel_grupo = ws.cell(row=i, column=1)
        cor_grupo = {"TROCAR": verde, "MANTER": None,
                     "SUBIU_NO_ATUAL": PatternFill("solid", fgColor="F8D7D7"),
                     "CONFERIR": amarelo, "SEM_PRECO": cinza}.get(r.get("grupo"))
        if cor_grupo is not None:
            cel_grupo.fill = cor_grupo
        cel_grupo.font = Font(bold=True)
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

    # Largura por CHAVE, nao por indice: acrescentar uma coluna no meio de
    # COLUNAS_SAIDA nao pode desalinhar as larguras de todas as seguintes.
    LARGURA_POR_CHAVE = {
        "grupo": 13, "fornecedor_atual": 26,
        "preco_fornecedor_atual_no_portal": 24, "variacao_no_fornecedor_atual": 20,
        "municipio_menor_preco": 20,
        "descricao_planilha": 42, "medida_planilha": 12,
        "estabelecimento_municipio": 30, "descricao_oferta": 40,
        "municipio_menor_preco": 24, "estabelecimento_menor_preco": 30,
        "endereco_menor_preco": 44, "gtin": 20, "data_venda": 22,
        "termo_consultado": 28, "observacao": 46,
        "preco_equivalente_na_medida_da_planilha": 22,
        "ofertas_descartadas_ruido": 22, "ofertas_descartadas_outlier": 24,
        "fornecedores_distintos": 20,
    }
    for n in (1, 2, 3):
        LARGURA_POR_CHAVE["alt%d_fornecedor" % n] = 30
        LARGURA_POR_CHAVE["alt%d_municipio" % n] = 18
        LARGURA_POR_CHAVE["alt%d_equivalente" % n] = 17
    larguras = {j: LARGURA_POR_CHAVE.get(chave, 16)
                for j, (chave, _r, _f) in enumerate(COLUNAS_SAIDA, 1)}
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


# --------------------------------------------------------------------------
#  PDF no padrao AIONS
# --------------------------------------------------------------------------
# Identidade herdada do relatorio da Esteira Financeira (mesma sessao, mesmo
# padrao): Montserrat para display e titulo, Lato para corpo e dado, tinta navy,
# acento teal so em regua e realce. Os arquivos de marca ficam em `marca/`:
# fontes reais e a logo em vetor, gerada do SVG de marca. Nada aqui redesenha a
# marca nem escolhe cor "parecida".
#
# Decisao medida, nao estetica: Montserrat tem digitos PROPORCIONAIS, entao toda
# coluna de valor usa Lato, cujos digitos sao tabulares. Numero em Montserrat
# aparece so isolado, no destaque de economia.

# tinta e acento, em RGB 0-1 (os mesmos hex do padrao)
AIONS_NAVY = (0.071, 0.118, 0.192)      # #121E31  tinta principal
AIONS_TEAL = (0.220, 0.639, 0.690)      # #38A3B0  acento da marca
AIONS_TEAL_DK = (0.173, 0.475, 0.533)   # #2C7988  acento sobre claro
AIONS_CLOUD = (0.973, 0.980, 0.988)     # #F8FAFC  superficie
AIONS_MINT = (0.941, 0.992, 0.980)      # #F0FDFA  superficie de destaque
AIONS_LINE = (0.886, 0.910, 0.941)      # #E2E8F0  regua
AIONS_SLATE = (0.392, 0.455, 0.545)     # #64748B  texto secundario
AIONS_SLATE_DK = (0.278, 0.333, 0.412)  # #475569  texto de tabela
AIONS_PERIGO = (0.718, 0.110, 0.110)    # #B71C1C  alta de preco, sempre com texto

_FONTES_OK = None


def _registrar_fontes():
    """
    Registra Montserrat e Lato de `marca/`. Devolve True se as quatro entraram.

    Fonte declarada e nao carregada e acabamento ruim: o documento cairia na
    fonte do sistema sem ninguem perceber. Por isso o resultado e explicito e o
    relatorio avisa quando saiu na fonte de reserva.
    """
    global _FONTES_OK
    if _FONTES_OK is not None:
        return _FONTES_OK
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    pasta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "marca")
    quatro = (("AIONS-Display", "Montserrat-Bold.ttf"),
              ("AIONS-Titulo", "Montserrat-SemiBold.ttf"),
              ("AIONS-Corpo", "Lato-Regular.ttf"),
              ("AIONS-Forte", "Lato-Bold.ttf"))
    try:
        for nome, arq in quatro:
            pdfmetrics.registerFont(TTFont(nome, os.path.join(pasta, arq)))
        _FONTES_OK = True
    except Exception:
        _FONTES_OK = False
    return _FONTES_OK


def escrever_pdf(resultados: list[dict], caminho: str, contexto: dict) -> None:
    """
    Relatorio de decisao no padrao AIONS, em quatro blocos:
    trocar, manter, subiu no proprio fornecedor, sem conclusao.

    Cada item de "trocar" e "manter" mostra a melhor opcao e, abaixo, as outras
    opcoes de fornecedor, para a decisao considerar distancia e nao so preco.
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate,
                                    Paragraph, Spacer, Table, TableStyle, KeepTogether)

    tem_fontes = _registrar_fontes()
    F_DISPLAY = "AIONS-Display" if tem_fontes else "Helvetica-Bold"
    F_TITULO = "AIONS-Titulo" if tem_fontes else "Helvetica-Bold"
    F_CORPO = "AIONS-Corpo" if tem_fontes else "Helvetica"
    F_FORTE = "AIONS-Forte" if tem_fontes else "Helvetica-Bold"

    NAVY = colors.Color(*AIONS_NAVY)
    TEAL = colors.Color(*AIONS_TEAL)
    TEAL_DK = colors.Color(*AIONS_TEAL_DK)
    CLOUD = colors.Color(*AIONS_CLOUD)
    MINT = colors.Color(*AIONS_MINT)
    LINE = colors.Color(*AIONS_LINE)
    SLATE = colors.Color(*AIONS_SLATE)
    SLATE_DK = colors.Color(*AIONS_SLATE_DK)
    PERIGO = colors.Color(*AIONS_PERIGO)

    # escala tipografica com poucos degraus, herdada do padrao
    T_H2, T_CORPO, T_TAB, T_MICRO, T_ALT = 11.5, 8.6, 8.2, 7.0, 7.4

    est = {
        "h2": ParagraphStyle("h2", fontName=F_TITULO, fontSize=T_H2, leading=14,
                             textColor=NAVY, spaceBefore=13, spaceAfter=2),
        "h2t": ParagraphStyle("h2t", fontName=F_TITULO, fontSize=T_H2, leading=15,
                              textColor=NAVY),
        # cabecalho de tabela: caixa alta com tracking, como na apresentacao
        "cab": ParagraphStyle("cab", fontName=F_FORTE, fontSize=T_MICRO,
                              leading=9, textColor=colors.white, charSpace=0.42),
        "cabd": ParagraphStyle("cabd", fontName=F_FORTE, fontSize=T_MICRO,
                               leading=9, textColor=colors.white, charSpace=0.42,
                               alignment=2),
        "nota": ParagraphStyle("nota", fontName=F_CORPO, fontSize=T_CORPO - 0.4,
                               leading=11.5, textColor=SLATE, spaceAfter=5),
        "cel": ParagraphStyle("cel", fontName=F_CORPO, fontSize=T_TAB, leading=9.6,
                              textColor=SLATE_DK),
        "celf": ParagraphStyle("celf", fontName=F_FORTE, fontSize=T_TAB, leading=9.6,
                               textColor=NAVY),
        "micro": ParagraphStyle("micro", fontName=F_CORPO, fontSize=T_MICRO,
                                leading=8.4, textColor=SLATE),
        # as linhas de "outra opção" ficavam em 6,6pt e mal se liam; 7,4 contra
        # 8,2 do corpo mantem a hierarquia e ainda e legivel no papel
        "alt": ParagraphStyle("alt", fontName=F_CORPO, fontSize=T_ALT,
                              leading=9.6, textColor=SLATE_DK),
        "fim": ParagraphStyle("fim", fontName=F_CORPO, fontSize=T_MICRO + 0.6,
                              leading=10.4, textColor=SLATE, spaceBefore=2.5),
    }

    MARG = 16 * mm
    doc = BaseDocTemplate(caminho, pagesize=A4, title="Radar de compras",
                          author="AIONS", leftMargin=MARG, rightMargin=MARG,
                          topMargin=MARG, bottomMargin=14 * mm)
    LARG = doc.width

    def masthead(canv, _doc):
        """
        Cabecalho conforme a APRESENTACAO PADRAO AIONS: logo no canto superior
        DIREITO, titulo navy, linha de contexto em teal, regua fina no rodape.
        Fundo branco, porque a logo tem faceta escura.
        """
        canv.saveState()
        alt = A4[1]
        # logo a direita, como em todas as paginas da apresentacao
        if _LOGO is not None:
            try:
                alt_logo = 15.5 * mm
                larg_logo = alt_logo * _LOGO.LARGURA / _LOGO.ALTURA
                _LOGO.desenhar(canv, MARG + LARG - larg_logo, 13 * mm, alt_logo)
            except Exception:
                pass
        canv.setFillColorRGB(*AIONS_NAVY)
        canv.setFont(F_DISPLAY, 19)
        canv.drawString(MARG, alt - 19 * mm, "Radar de compras")
        # o subtitulo da apresentacao e teal, nao cinza
        canv.setFont(F_FORTE, T_CORPO)
        canv.setFillColorRGB(*AIONS_TEAL_DK)
        canv.drawString(MARG, alt - 26.5 * mm,
                        "%s, %s" % (contexto.get("municipio", ""), contexto.get("data", "")))
        canv.setFont(F_CORPO, T_CORPO - 0.4)
        canv.setFillColorRGB(*AIONS_SLATE)
        canv.drawString(MARG, alt - 32 * mm,
                        "%s, %d itens conferidos" % (contexto.get("portal", ""),
                                                     len(resultados)))
        canv.setStrokeColorRGB(*AIONS_TEAL)
        canv.setLineWidth(2.4)
        canv.line(MARG, alt - 37 * mm, MARG + LARG, alt - 37 * mm)
        # regua de fechamento no pe da pagina, como na apresentacao
        canv.setStrokeColorRGB(*AIONS_LINE)
        canv.setLineWidth(0.6)
        canv.line(MARG, 12.5 * mm, MARG + LARG, 12.5 * mm)
        canv.setFont(F_CORPO, T_MICRO)
        canv.setFillColorRGB(*AIONS_SLATE)
        canv.drawRightString(MARG + LARG, 9 * mm, "%d" % _doc.page)
        canv.restoreState()

    quadro = Frame(MARG, 14 * mm, LARG, A4[1] - 40 * mm - 14 * mm, id="corpo",
                   leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="aions", frames=[quadro], onPage=masthead)])

    def brl(v):
        if v is None:
            return "-"
        # o sinal vem ANTES do simbolo: "-R$ 0,24", nunca "R$ -0,24"
        return ("-" if v < 0 else "") + ("R$ %.2f" % abs(v)).replace(".", ",")

    def km(v):
        return "n/d" if v is None else ("%.1f" % v).replace(".", ",")

    por_grupo = {}
    for r in resultados:
        por_grupo.setdefault(r.get("grupo", "SEM_PRECO"), []).append(r)
    trocar = sorted(por_grupo.get("TROCAR", []), key=lambda r: -(r["economia_unitaria"] or 0))
    manter = por_grupo.get("MANTER", [])
    subiu = sorted(por_grupo.get("SUBIU_NO_ATUAL", []),
                   key=lambda r: -(r["variacao_no_fornecedor_atual"] or 0))
    pendentes = por_grupo.get("CONFERIR", []) + por_grupo.get("SEM_PRECO", [])
    total = sum(r["economia_unitaria"] for r in trocar)

    hist = []

    # ---- destaque: o numero que resume a leitura -------------------------
    if trocar:
        hist.append(Spacer(1, 5))          # respiro sob a regua de acento
        painel = Table(
            [[Paragraph('<font name="%s" size="26" color="#121E31">%s</font><br/>'
                        '<br/><font name="%s" size="7.4" color="#64748B">economia por '
                        'unidade, somando %d %s</font>'
                        % (F_DISPLAY, brl(total), F_CORPO, len(trocar),
                           "item" if len(trocar) == 1 else "itens"),
                        ParagraphStyle("destaque", fontName=F_CORPO, fontSize=8,
                                       leading=15, textColor=SLATE_DK)),
              Paragraph('<font name="%s" size="7.6" color="#475569">'
                        'Cada preço abaixo é o equivalente à embalagem da sua planilha. '
                        'Os valores vêm de NFC-e já emitida: é o que alguém pagou, '
                        'não oferta vigente.</font>' % F_CORPO, est["cel"])]],
            colWidths=[LARG * 0.36, LARG * 0.64])
        painel.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), MINT),
            ("BACKGROUND", (1, 0), (1, 0), CLOUD),
            ("LINEBEFORE", (0, 0), (0, 0), 2.4, TEAL),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ("TOPPADDING", (0, 0), (-1, -1), 12),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ]))
        hist.append(painel)

    def titulo_secao(texto):
        """
        Titulo com marcador vertical teal a esquerda: e a assinatura de secao da
        APRESENTACAO PADRAO AIONS, e o unico lugar, alem da regua, onde o acento
        aparece em area.
        """
        t = Table([["", Paragraph(texto, est["h2t"])]],
                  colWidths=[2.6, LARG - 2.6])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), TEAL),
            ("LEFTPADDING", (0, 0), (0, 0), 0), ("RIGHTPADDING", (0, 0), (0, 0), 0),
            ("LEFTPADDING", (1, 0), (1, 0), 7), ("RIGHTPADDING", (1, 0), (1, 0), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        return t

    def estilo_tabela(n_linhas, destaque=None, principais=()):
        e = [
            ("BACKGROUND", (0, 0), (-1, 0), NAVY),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), F_FORTE),
            ("FONTSIZE", (0, 0), (-1, 0), T_MICRO + 0.4),
            ("FONTNAME", (0, 1), (-1, -1), F_CORPO),
            ("FONTSIZE", (0, 1), (-1, -1), T_TAB),
            ("TEXTCOLOR", (0, 1), (-1, -1), SLATE_DK),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
            ("LINEBELOW", (0, 0), (-1, -1), 0.5, LINE),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ]
        for i in principais:                      # linha da melhor opcao do item
            e.append(("BACKGROUND", (0, i), (-1, i), MINT))
            e.append(("FONTNAME", (0, i), (-1, i), F_FORTE))
            e.append(("TEXTCOLOR", (0, i), (-1, i), NAVY))
        if destaque is not None:
            e.append(("TEXTCOLOR", (-1, 1), (-1, -1), destaque))
            e.append(("FONTNAME", (-1, 1), (-1, -1), F_FORTE))
        return TableStyle(e)

    # ---- 1. trocar, com as outras opcoes de cada item -------------------
    hist.append(Spacer(1, 12))
    hist.append(titulo_secao("1. Vale trocar de fornecedor: %d %s"
                             % (len(trocar), "item" if len(trocar) == 1 else "itens")))
    hist.append(Spacer(1, 3))
    if trocar:
        hist.append(Paragraph("A linha destacada é a melhor opção. Abaixo dela, as outras "
                              "opções do mesmo item, ordenadas por preço na sua embalagem.",
                              est["nota"]))
        COLS = (0.28, 0.105, 0.095, 0.20, 0.145, 0.055, 0.12)
        CABECALHO = [Paragraph("PRODUTO E OPÇÕES", est["cab"]),
                     Paragraph("VOCÊ PAGA", est["cabd"]),
                     Paragraph("PREÇO", est["cabd"]),
                     Paragraph("FORNECEDOR", est["cab"]),
                     Paragraph("MUNICÍPIO", est["cab"]),
                     Paragraph("KM", est["cabd"]),
                     Paragraph("ECONOMIA", est["cabd"])]

        def cabecalho():
            t = Table([CABECALHO], colWidths=[LARG * x for x in COLS])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), F_FORTE),
                ("FONTSIZE", (0, 0), (-1, 0), T_MICRO + 0.4),
                ("ALIGN", (1, 0), (2, 0), "RIGHT"),     # Você paga, Preço
                ("ALIGN", (3, 0), (4, 0), "LEFT"),      # Fornecedor, Município
                ("ALIGN", (5, 0), (-1, 0), "RIGHT"),    # km, Economia
                ("TOPPADDING", (0, 0), (-1, 0), 3.5),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 3.5),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ]))
            return t

        # UMA tabela por item, dentro de KeepTogether: o item e suas opcoes nunca
        # se separam. Numa tabela unica, o reportlab quebra em qualquer linha, e
        # as alternativas apareciam orfas no topo da pagina seguinte, sem o
        # produto a que pertencem.
        hist.append(cabecalho())
        for r in trocar:
            linhas_item = [[
                Paragraph(str(r["descricao_planilha"])[:44], est["celf"]),
                brl(r["preco_atual"]),
                brl(r.get("preco_equivalente_na_medida_da_planilha")),
                Paragraph(str(r["estabelecimento_menor_preco"])[:26], est["celf"]),
                Paragraph(str(r["municipio_menor_preco"] or "-")[:18], est["celf"]),
                km(r.get("distancia_km")),
                brl(r["economia_unitaria"]),
            ]]
            comparaveis_alt = [a for a in (r.get("alternativas") or [])[1:]
                               if a.get("preco_na_medida_da_planilha") is not None]
            for n, alt in enumerate(comparaveis_alt[:3], 1):
                ordinal = {1: "2ª", 2: "3ª", 3: "4ª"}[n]
                linhas_item.append([
                    Paragraph('<font name="%s" color="#475569">%s opção</font>'
                              '&nbsp;<font color="#64748B">%s</font>'
                              % (F_FORTE, ordinal, str(alt.get("embalagem") or "")),
                              est["alt"]),
                    "",
                    Paragraph(brl(alt.get("preco_na_medida_da_planilha")), est["alt"]),
                    Paragraph(str(alt.get("fornecedor") or "-")[:26], est["alt"]),
                    Paragraph(str(alt.get("municipio") or "-")[:18], est["alt"]),
                    Paragraph(km(alt.get("distancia_km")), est["alt"]),
                    Paragraph(brl(alt.get("economia_vs_atual")), est["alt"]),
                ])
            t = Table(linhas_item, colWidths=[LARG * x for x in COLS])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), MINT),
                ("FONTNAME", (0, 0), (-1, 0), F_FORTE),
                ("TEXTCOLOR", (0, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (-1, 0), (-1, 0), TEAL_DK),
                ("FONTNAME", (0, 1), (-1, -1), F_CORPO),
                ("FONTSIZE", (0, 0), (-1, -1), T_TAB),
                ("TEXTCOLOR", (0, 1), (-1, -1), SLATE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (2, -1), "RIGHT"),
                ("ALIGN", (5, 0), (-1, -1), "RIGHT"),
                ("LINEBELOW", (0, 0), (-1, -1), 0.5, LINE),
                ("TOPPADDING", (0, 0), (-1, -1), 3.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ]))
            hist.append(KeepTogether(t))
    else:
        hist.append(Paragraph("Nenhum item está mais barato em outro fornecedor.", est["nota"]))

    # ---- 2. manter -------------------------------------------------------
    hist.append(Spacer(1, 13))
    hist.append(titulo_secao("2. Vale manter o fornecedor atual: %d %s"
                             % (len(manter), "item" if len(manter) == 1 else "itens")))
    hist.append(Spacer(1, 3))
    if manter:
        hist.append(Paragraph("O melhor preço do mercado é igual ou maior do que você já paga.",
                              est["nota"]))
        dados = [[Paragraph("PRODUTO", est["cab"]), Paragraph("VOCÊ PAGA", est["cabd"]), Paragraph("MELHOR DO MERCADO", est["cabd"]), Paragraph("FORNECEDOR", est["cab"]), Paragraph("MUNICÍPIO", est["cab"]), Paragraph("DIFERENÇA", est["cabd"])]]
        for r in sorted(manter, key=lambda r: r["economia_unitaria"] or 0):
            dados.append([
                Paragraph(str(r["descricao_planilha"])[:44], est["celf"]),
                brl(r["preco_atual"]),
                brl(r.get("preco_equivalente_na_medida_da_planilha") or r["menor_preco_estado"]),
                Paragraph(str(r["estabelecimento_menor_preco"])[:26], est["cel"]),
                Paragraph(str(r["municipio_menor_preco"] or "-")[:18], est["cel"]),
                brl(r["economia_unitaria"]),
            ])
        t = Table(dados, colWidths=[LARG * x for x in (0.28, 0.10, 0.15, 0.21, 0.16, 0.10)],
                  repeatRows=1)
        t.setStyle(estilo_tabela(len(dados), SLATE))
        hist.append(t)
    else:
        hist.append(Paragraph("Nenhum item nessa situação.", est["nota"]))

    # ---- 3. subiu no proprio fornecedor ---------------------------------
    hist.append(Spacer(1, 13))
    hist.append(titulo_secao("3. Preço subiu no seu próprio fornecedor: %d %s"
                             % (len(subiu), "item" if len(subiu) == 1 else "itens")))
    hist.append(Spacer(1, 3))
    if subiu:
        hist.append(Paragraph("O seu fornecedor aparece no portal cobrando mais do que o valor "
                              "registrado na sua planilha. Confira se o preço combinado "
                              "ainda vale.", est["nota"]))
        dados = [[Paragraph("PRODUTO", est["cab"]), Paragraph("SUA PLANILHA", est["cabd"]), Paragraph("MESMO FORNECEDOR HOJE", est["cabd"]), Paragraph("FORNECEDOR", est["cab"]), Paragraph("MUNICÍPIO", est["cab"]), Paragraph("ALTA", est["cabd"])]]
        for r in subiu:
            dados.append([
                Paragraph(str(r["descricao_planilha"])[:44], est["celf"]),
                brl(r["preco_atual"]),
                brl(r["preco_fornecedor_atual_no_portal"]),
                Paragraph(str(r.get("fornecedor_atual") or "-")[:26], est["cel"]),
                Paragraph(str(r["municipio_menor_preco"] or "-")[:18], est["cel"]),
                "+" + brl(r["variacao_no_fornecedor_atual"]),
            ])
        t = Table(dados, colWidths=[LARG * x for x in (0.27, 0.11, 0.17, 0.22, 0.14, 0.09)],
                  repeatRows=1)
        t.setStyle(estilo_tabela(len(dados), PERIGO))
        hist.append(t)
    else:
        hist.append(Paragraph("Nenhum item nessa situação, o que só vale para os fornecedores "
                              "da sua planilha que aparecem no portal. Quem não emitiu NFC-e "
                              "no período não pode ser verificado.", est["nota"]))

    # ---- 4. sem conclusao ------------------------------------------------
    if pendentes:
        hist.append(Spacer(1, 13))
        hist.append(titulo_secao("4. Sem conclusão: %d %s"
                                 % (len(pendentes), "item" if len(pendentes) == 1 else "itens")))
        hist.append(Spacer(1, 3))
        dados = [[Paragraph("PRODUTO", est["cab"]), Paragraph("VOCÊ PAGA", est["cabd"]), Paragraph("POR QUÊ", est["cab"])]]
        for r in pendentes:
            if r["menor_preco_estado"] is None:
                motivo = "nenhuma venda deste produto no período consultado"
            else:
                partes = [t.strip() for t in str(r.get("observacao") or "").split("|")]
                motivo = partes[-1] if (partes and partes[-1]) else (
                    "sem base para comparar com a embalagem da sua planilha")
            dados.append([Paragraph(str(r["descricao_planilha"])[:44], est["celf"]),
                          brl(r["preco_atual"]),
                          Paragraph(motivo.replace("--", ":"), est["micro"])])
        t = Table(dados, colWidths=[LARG * x for x in (0.27, 0.10, 0.63)], repeatRows=1)
        t.setStyle(estilo_tabela(len(dados)))
        hist.append(t)

    # ---- ressalvas: o minimo que muda a decisao -------------------------
    ruido = sum(r.get("ofertas_descartadas_ruido") or 0 for r in resultados)
    fora = sum(r.get("ofertas_descartadas_outlier") or 0 for r in resultados)
    hist.append(Spacer(1, 11))
    hist.append(titulo_secao("Antes de decidir"))
    hist.append(Spacer(1, 3))
    linhas = [
        "O estabelecimento não é obrigado a manter o preço, e o produto pode estar sem "
        "estoque.",
        "A economia é por unidade. Multiplique pelas quantidades que você compra.",
        "Não inclui frete nem diferença de ICMS. Distância em linha reta." +
        (" No Amazonas, confirme o acesso: muitos municípios só têm ligação fluvial."
         if contexto.get("uf") == "AM" else ""),
    ]
    if ruido or fora:
        linhas.append("Descartados antes das contas: %d resultado(s) de outro produto e %d "
                      "com preço fora da distribuição, como nota de brinde ou erro de "
                      "digitação. Detalhe na aba Auditoria da planilha." % (ruido, fora))
    if not tem_fontes:
        linhas.append("Este PDF saiu na fonte de reserva: os arquivos de marca não foram "
                      "encontrados em marca/.")
    linhas.append("Este relatório não recomenda trocar de fornecedor. Ele mostra preço, "
                  "distância e origem lado a lado.")
    for t in linhas:
        hist.append(Paragraph(t, est["fim"]))

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
    check("so BAIXA -> sinalizado na observacao",
          "casamento fraco" in r_baixa["observacao"], f"-> {r_baixa['observacao'][:50]}")

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

    print("descricoes torpes do portal (vistas em 09/09/2026)")
    # virgula decimal perdida na digitacao
    m_esp = extrair_medida("SUCO DEL VALLE UVA 1 5L")
    check("'1 5L' e 1,5 litro, nao 5 litros",
          m_esp.total == 1500.0 and m_esp.ambigua, f"-> {m_esp}")
    check("'1 5L' nao vira 5000 ml", extrair_medida("SUCO DEL VALLE UVA 1 5L").total != 5000.0)
    check("virgula normal continua funcionando",
          extrair_medida("REFRIGERANTE COCA COLA 1,5L").total == 1500.0)
    check("multipack legitimo nao e confundido com decimal perdido",
          extrair_medida("CERVEJA 12X500ML").total == 6000.0)

    # contagem de unidades
    check("'OVOS BRANCOS 06 UN' -> 6 unidades",
          (extrair_medida("OVOS BRANCOS 06 UN").total,
           extrair_medida("OVOS BRANCOS 06 UN").base) == (6.0, "un"))
    check("'DUZIA' -> 12 unidades",
          (extrair_medida("OVOS BRANCOS DUZIA").total,
           extrair_medida("OVOS BRANCOS DUZIA").base) == (12.0, "un"))
    check("'MEIA DUZIA' -> 6 unidades",
          extrair_medida("OVOS MEIA DUZIA").total == 6.0)
    check("volume tem prioridade sobre contagem",
          extrair_medida("REFRIGERANTE 2L 6 UN").base == "ml",
          f"-> {extrair_medida('REFRIGERANTE 2L 6 UN')}")

    # o caso real: duzia contra meia duzia
    r_ovos = consolidar("OVOS BRANCOS DUZIA", 12.90,
                        [Oferta("OVOS BRANCOS 06 UN", 4.25, "MERCADINHO", "", "Manaus",
                                medida=extrair_medida("OVOS BRANCOS 06 UN"))],
                        "Manaus", "AM", None, geocode=False)
    check("duzia x meia duzia: equivalente e o dobro, nao a etiqueta",
          abs(r_ovos["preco_equivalente_na_medida_da_planilha"] - 8.50) < 0.01,
          f"-> {r_ovos['preco_equivalente_na_medida_da_planilha']}")
    check("duzia x meia duzia: economia real, nao 67%",
          abs(r_ovos["economia_unitaria"] - 4.40) < 0.01,
          f"-> {r_ovos['economia_unitaria']}")

    # oferta sem medida quando a planilha tem: nao inventa economia
    r_semmed = consolidar("LINGUICA PERDIGAO 500G", 24.90,
                          [Oferta("PERDIGAO LINGUICA MI", 8.75, "VAREJAO", "", "Manaus",
                                  medida=extrair_medida("PERDIGAO LINGUICA MI"))],
                          "Manaus", "AM", None, geocode=False)
    check("oferta sem gramatura -> economia NAO calculada",
          r_semmed["economia_unitaria"] is None, f"-> {r_semmed['economia_unitaria']}")
    check("oferta sem gramatura -> motivo declarado",
          "informa o tamanho" in r_semmed["observacao"],
          f"-> {r_semmed['observacao'][:60]}")
    check("mas o preco fica visivel como referencia",
          r_semmed["menor_preco_estado"] == 8.75)

    # planilha tambem sem medida: comparacao absoluta segue valendo
    r_ambos = consolidar("CEBOLA GRANEL", 4.99,
                         [Oferta("CEBOLA GRANEL", 3.99, "FEIRA", "", "Manaus",
                                 medida=extrair_medida("CEBOLA GRANEL"))],
                         "Manaus", "AM", None, geocode=False)
    check("sem medida nos dois lados -> compara absoluto",
          abs((r_ambos["economia_unitaria"] or 0) - 1.00) < 0.01,
          f"-> {r_ambos['economia_unitaria']}")

    print("alternativas achatadas nas colunas (Alternativa 1/2/3)")
    lojas_alt = [
        ("MERCADO A", 5.99, "Manaus"), ("MERCADO B", 6.19, "Manaus"),
        ("MERCADO C", 6.49, "Itacoatiara"), ("MERCADO D", 6.99, "Manaus"),
        ("MERCADO E", 7.49, "Manaus"),
    ]
    r_alt = consolidar("ARROZ TIO JOAO 1KG", 8.90,
                       [Oferta("ARROZ TIO JOAO 1KG", pr, nome, "", mun,
                               medida=extrair_medida("ARROZ TIO JOAO 1KG"))
                        for nome, pr, mun in lojas_alt],
                       "Manaus", "AM", None, geocode=False)
    check("menor no estado e a 1a opcao", r_alt["menor_preco_estado"] == 5.99)
    check("Alternativa 1 e a SEGUNDA melhor, nao a primeira",
          r_alt["alt1_preco"] == 6.19 and r_alt["alt1_fornecedor"] == "MERCADO B",
          f"-> {r_alt['alt1_preco']} / {r_alt['alt1_fornecedor']}")
    check("Alternativa 2 e 3 seguem a ordem de preco",
          (r_alt["alt2_preco"], r_alt["alt3_preco"]) == (6.49, 6.99),
          f"-> {r_alt['alt2_preco']}, {r_alt['alt3_preco']}")
    check("preco e fornecedor da alternativa nao se cruzam",
          r_alt["alt2_fornecedor"] == "MERCADO C" and r_alt["alt2_municipio"] == "Itacoatiara")
    check("nunca repete o menor preco como alternativa",
          r_alt["alt1_fornecedor"] != r_alt["estabelecimento_menor_preco"])
    r_poucas = consolidar("ARROZ TIO JOAO 1KG", 8.90,
                          [Oferta("ARROZ TIO JOAO 1KG", 5.99, "UNICO", "", "Manaus",
                                  medida=extrair_medida("ARROZ TIO JOAO 1KG"))],
                          "Manaus", "AM", None, geocode=False)
    check("uma oferta so -> alternativas em branco",
          r_poucas["alt1_preco"] is None and r_poucas["alt1_fornecedor"] == "")
    check("item sem oferta -> alternativas em branco",
          consolidar("CEBOLA", 4.99, [], "Manaus", "AM", None,
                     geocode=False)["alt3_preco"] is None)
    # a ordem e por preco na medida da planilha, nao por preco de etiqueta:
    # um fardo de 12 L a R$ 58,00 e mais barato POR LITRO que uma garrafa a
    # R$ 2,79, e por isso aparece antes mesmo custando 20x mais
    lojas_emb = [
        ("FARDO", 58.00, "DETERGENTE YPE NEUTRO 24 X 500ML"),
        ("GARRAFA A", 2.79, "DETERGENTE YPE NEUTRO 500ML"),
        ("GARRAFA B", 2.89, "DETERGENTE YPE NEUTRO 500ML"),
        ("GARRAFA C", 2.99, "DETERGENTE YPE NEUTRO 500ML"),
    ]
    r_emb = consolidar("DETERGENTE YPE NEUTRO 500ML", 3.49,
                       [Oferta(desc, pr, nome, "", "Manaus", medida=extrair_medida(desc))
                        for nome, pr, desc in lojas_emb],
                       "Manaus", "AM", None, geocode=False)
    equivalentes = [r_emb["preco_equivalente_na_medida_da_planilha"]] + [
        r_emb["alt%d_equivalente" % n] for n in (1, 2, 3)]
    equivalentes = [e for e in equivalentes if e is not None]
    check("equivalentes ficam em ordem crescente (a ordem real)",
          equivalentes == sorted(equivalentes), f"-> {equivalentes}")
    check("o fardo caro por etiqueta e o melhor por litro",
          r_emb["menor_preco_estado"] == 58.00
          and r_emb["preco_equivalente_na_medida_da_planilha"] < r_emb["alt1_equivalente"],
          f"-> {r_emb['menor_preco_estado']} / {r_emb['preco_equivalente_na_medida_da_planilha']}")

    # a aba Comparativo leva 4 campos por alternativa; o preco de etiqueta fica
    # na aba Alternativas, para a principal nao virar um paredao de colunas
    chaves_col = [k for k, _r, _f in COLUNAS_SAIDA]
    check("cada alternativa tem preco, fornecedor, municipio e km na planilha",
          all(("alt%d_%s" % (n, c)) in chaves_col
              for n in (1, 2, 3)
              for c in ("equivalente", "fornecedor", "municipio", "km")))
    check("todo preco da planilha tem municipio ao lado",
          all(k in chaves_col for k in ("municipio_menor_preco", "alt1_municipio",
                                        "alt2_municipio", "alt3_municipio")))
    check("alt_preco continua no dado, mesmo fora da aba principal",
          "alt1_preco" in r_alt)

    print("gera os arquivos de saida de verdade (pega erro de programacao)")
    import os as _os, tempfile as _tmp
    linhas_falsas = [
        consolidar("ARROZ 1KG", 8.90,
                   [Oferta("ARROZ TIO JOAO 1KG", 5.99, "MERCADO X", "RUA A, MANAUS-AM",
                           "Manaus", medida=extrair_medida("ARROZ 1KG"), gtin="7893500020127"),
                    Oferta("ARROZ TIO JOAO 1KG", 6.49, "MERCADO Y", "RUA B, MANAUS-AM",
                           "Manaus", medida=extrair_medida("ARROZ 1KG"))],
                   "Manaus", "AM", None, geocode=False, fornecedor_atual="MERCADO Y"),
        consolidar("CEBOLA", 4.99, [], "Manaus", "AM", None, geocode=False),
    ]
    ctx_falso = {"uf": "AM", "municipio": "Manaus", "portal": "Busca Preco AM",
                 "data": "01/01/2026 00:00"}
    base = _os.path.join(_tmp.mkdtemp(), "saida")
    erro_xlsx = erro_pdf = None
    try:
        escrever_xlsx(linhas_falsas, base + ".xlsx", ctx_falso)
    except Exception as e:
        erro_xlsx = repr(e)
    try:
        escrever_pdf(linhas_falsas, base + ".pdf", ctx_falso)
    except Exception as e:
        erro_pdf = repr(e)
    check("escrever_xlsx roda sem excecao", erro_xlsx is None, f"-> {erro_xlsx}")
    check("escrever_pdf roda sem excecao", erro_pdf is None, f"-> {erro_pdf}")
    check("o .xlsx existe e nao esta vazio",
          _os.path.exists(base + ".xlsx") and _os.path.getsize(base + ".xlsx") > 4000)
    check("o .pdf existe e nao esta vazio",
          _os.path.exists(base + ".pdf") and _os.path.getsize(base + ".pdf") > 1500)
    if erro_xlsx is None:
        from openpyxl import load_workbook as _lw
        _wb = _lw(base + ".xlsx")
        check("as 4 abas foram criadas",
              _wb.sheetnames == ["Comparativo", "Resumo", "Alternativas", "Auditoria"],
              f"-> {_wb.sheetnames}")
        _cab = [c.value for c in _wb["Comparativo"][1]]
        check("cabecalho da planilha bate com COLUNAS_SAIDA",
              len(_cab) == len(COLUNAS_SAIDA) and _cab[0] == COLUNAS_SAIDA[0][1])
        check("a coluna de decisao sai com rotulo legivel",
              str(_wb["Comparativo"].cell(row=2, column=1).value or "").split()[0] in
              ("1", "2", "3", "4"),
              f"-> {_wb['Comparativo'].cell(row=2, column=1).value}")

    print("grupos de decisao")
    r_troca = consolidar("ARROZ 1KG", 8.90,
                         [Oferta("ARROZ 1KG", 5.99, "MERCADO X", "", "Manaus",
                                 medida=extrair_medida("ARROZ 1KG"))],
                         "Manaus", "AM", None, geocode=False, fornecedor_atual="Rio Negro")
    check("mercado mais barato -> TROCAR", r_troca["grupo"] == "TROCAR", f"-> {r_troca['grupo']}")
    r_mant = consolidar("ARROZ 1KG", 5.00,
                        [Oferta("ARROZ 1KG", 5.99, "MERCADO X", "", "Manaus",
                                medida=extrair_medida("ARROZ 1KG"))],
                        "Manaus", "AM", None, geocode=False, fornecedor_atual="Rio Negro")
    check("voce ja paga menos -> MANTER", r_mant["grupo"] == "MANTER", f"-> {r_mant['grupo']}")
    # o proprio fornecedor da planilha aparece no portal cobrando mais
    r_subiu = consolidar("ARROZ 1KG", 5.00,
                         [Oferta("ARROZ 1KG", 6.50, "RIO NEGRO ALIMENTOS LTDA", "", "Manaus",
                                 medida=extrair_medida("ARROZ 1KG"))],
                         "Manaus", "AM", None, geocode=False,
                         fornecedor_atual="Distribuidora Rio Negro")
    check("mesmo fornecedor cobrando mais -> SUBIU_NO_ATUAL",
          r_subiu["grupo"] == "SUBIU_NO_ATUAL", f"-> {r_subiu['grupo']}")
    check("a alta e quantificada", abs(r_subiu["variacao_no_fornecedor_atual"] - 1.50) < 0.01,
          f"-> {r_subiu['variacao_no_fornecedor_atual']}")
    r_semp = consolidar("CEBOLA", 4.99, [], "Manaus", "AM", None, geocode=False)
    check("sem oferta -> SEM_PRECO", r_semp["grupo"] == "SEM_PRECO")

    print("casamento de nome de fornecedor")
    check("razao social casa com nome curto",
          mesmo_fornecedor("Higiluz Comercial", "HIGILUZ COMERCIO DE PRODUTOS LTDA"))
    check("nao casa por palavra genérica",
          not mesmo_fornecedor("Distribuidora Rio Negro", "COMERCIAL SOARES LTDA"))
    check("nao casa fornecedores diferentes",
          not mesmo_fornecedor("Bebidas Amazonas", "MERCADINHO CEZAR"))
    check("uma palavra distintiva basta",
          mesmo_fornecedor("Atacado Ponta Negra", "PONTA NEGRA COMERCIO DE ALIMENTOS"))

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
          "fora da distribui" in r_out["observacao"], f"-> {r_out['observacao'][:70]}")
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
        forn = str(linha.get(cols["fornecedor"]) or "").strip() if cols.get("fornecedor") else ""
        resultados.append(
            consolidar(desc, preco, ofertas, municipio, args.uf, coord_base,
                       fornecedor_atual=forn)
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
        # o except existe para nao perder o PDF se o xlsx falhar, mas sem o
        # traceback ele esconde erro de programacao: uma variavel usada antes
        # de existir deixou de gerar a planilha por uma rodada inteira, em
        # silencio. Falha de saida e ALTO, nao um aviso discreto.
        import traceback
        print("=" * 62, file=sys.stderr)
        print("ERRO: a planilha NAO foi gravada -- %s" % erro, file=sys.stderr)
        traceback.print_exc()
        print("=" * 62, file=sys.stderr)
    try:
        escrever_pdf(resultados, base_saida + ".pdf", contexto)
        print(f"pdf em {base_saida}.pdf")
    except Exception as erro:
        import traceback
        print("=" * 62, file=sys.stderr)
        print("ERRO: o PDF NAO foi gravado -- %s" % erro, file=sys.stderr)
        traceback.print_exc()
        print("=" * 62, file=sys.stderr)
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
