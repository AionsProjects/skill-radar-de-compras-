#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Radar de Compras como servidor MCP, para o Claude Desktop.

POR QUE ISTO EXISTE
-------------------
A mesma logica ja roda como skill do Claude Code. No Claude Desktop, porem,
skill executa no ambiente de codigo da Anthropic, que nao alcanca o portal da
SEFAZ -- a ferramenta apareceria na lista e nao conseguiria consultar nada.

Servidor MCP roda na MAQUINA do usuario, com o Python e a rede dele. E o unico
jeito de o Desktop fazer a coleta de verdade.

Nao reimplementa nada: importa busca_preco.py, que continua sendo a fonte
unica da logica, dos filtros e do relatorio.

    python mcp_server.py            # stdio, que e o que o Desktop fala
"""

from __future__ import annotations

import io
import json
import os
import sys
import contextlib
from pathlib import Path
from typing import Any

# O Desktop inicia este processo com o diretorio de trabalho dele, nao com o
# nosso: sem isto o import de busca_preco falha dependendo de onde ele subiu.
AQUI = Path(__file__).resolve().parent
if str(AQUI) not in sys.path:
    sys.path.insert(0, str(AQUI))

from mcp.server.mcpserver import MCPServer  # noqa: E402
import busca_preco as bp                    # noqa: E402

mcp = MCPServer(
    name="radar-de-compras",
    instructions=(
        "Compara precos de uma planilha com o que o comercio praticou de fato, "
        "pelos portais oficiais de NFC-e do Amazonas e da Paraiba. Os valores "
        "vem de nota fiscal emitida: e o que alguem pagou, nao oferta anunciada. "
        "Nunca emite veredito de 'vale a pena' -- reporta economia e distancia "
        "lado a lado, e declara o que nao pode confirmar."
    ),
)

UFS = ("AM", "PB")
PADRAO_MUNICIPIO = {"AM": "Manaus", "PB": "Joao Pessoa"}


def _erro(msg: str, **extra: Any) -> dict:
    return {"ok": False, "erro": msg, **extra}


def _checar_uf(uf: str) -> str | None:
    if uf.upper() not in UFS:
        return None
    return uf.upper()


@mcp.tool(
    title="Cotar planilha",
    description=(
        "Le uma planilha (.xlsx/.csv) de produtos e compara com os precos "
        "praticados no comercio, gerando PDF e planilha anotada. A planilha "
        "precisa de uma coluna de produto; a de preco e opcional (sem ela o "
        "relatorio vira uma cotacao de mercado, sem economia); a de fornecedor "
        "habilita a analise do proprio fornecedor."
    ),
)
def cotar_planilha(
    caminho_planilha: str,
    uf: str = "AM",
    municipio: str = "",
    pasta_saida: str = "",
) -> dict:
    """Fluxo completo: le, consulta o portal, grava PDF + XLSX + JSON."""
    uf_ok = _checar_uf(uf)
    if not uf_ok:
        return _erro("UF nao suportada: %r. Esta ferramenta cobre %s."
                     % (uf, " e ".join(UFS)))
    planilha = Path(caminho_planilha).expanduser()
    if not planilha.is_file():
        return _erro("nao encontrei a planilha", caminho=str(planilha))

    municipio = municipio or PADRAO_MUNICIPIO[uf_ok]
    destino = Path(pasta_saida).expanduser() if pasta_saida else planilha.parent
    destino.mkdir(parents=True, exist_ok=True)
    saida = destino / ("radar-%s" % planilha.stem)

    argv = ["--uf", uf_ok, "--municipio", municipio,
            "--planilha", str(planilha), "--saida", str(saida)]

    # O protocolo MCP fala JSON-RPC pelo stdout: qualquer print do motor
    # corromperia a conversa com o Desktop. Por isso a saida e capturada e
    # devolvida como dado, nao impressa.
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            codigo = bp.main(argv)
    except Exception as e:                     # falha de rede, portal fora etc.
        return _erro("a consulta falhou: %s" % e, log=buf.getvalue()[-1500:])
    if codigo != 0:
        return _erro("o comparador terminou com erro", log=buf.getvalue()[-1500:])

    arquivos = {ext: str(saida.with_suffix("." + ext))
                for ext in ("pdf", "xlsx", "json")
                if saida.with_suffix("." + ext).is_file()}

    resumo: dict[str, Any] = {}
    caminho_json = saida.with_suffix(".json")
    if caminho_json.is_file():
        dados = json.loads(caminho_json.read_text(encoding="utf-8"))
        grupos: dict[str, int] = {}
        for r in dados:
            grupos[r.get("grupo", "?")] = grupos.get(r.get("grupo", "?"), 0) + 1
        resumo = {
            "itens": len(dados),
            "mais_barato_no_mercado": grupos.get("TROCAR", 0),
            "vale_manter": grupos.get("MANTER", 0),
            "so_cotacao": grupos.get("COTACAO", 0),
            "sem_conclusao": grupos.get("CONFERIR", 0) + grupos.get("SEM_PRECO", 0),
            # quem subiu no mercado nao e um grupo: e medido por item
            "subiu_no_mercado": sum(
                1 for r in dados
                if (r.get("alta_no_mercado") or 0) > 0
                and (r.get("mercado_amostra") or 0) >= 3
                and (r.get("alta_no_mercado_pct") or 0) >= 2.0
                and r.get("medida_planilha")
                and r.get("medida_planilha") == r.get("medida_oferta")
            ),
        }

    return {
        "ok": True,
        "municipio": municipio,
        "uf": uf_ok,
        "arquivos": arquivos,
        "resumo": resumo,
        "leia_o_pdf": arquivos.get("pdf", ""),
        "ressalvas": [
            "Precos vem de NFC-e ja emitida: e o que alguem pagou, nao oferta vigente.",
            "Nao inclui frete nem diferenca de ICMS. Distancia em linha reta.",
            "O estabelecimento nao e obrigado a manter o preco.",
        ],
    }


@mcp.tool(
    title="Consultar um produto",
    description=(
        "Consulta avulsa de um produto no portal, sem planilha. Devolve as "
        "ofertas mais baratas com loja, municipio e data da venda. Util para "
        "conferir um item so ou testar se o portal esta respondendo."
    ),
)
def consultar_produto(
    termo: str,
    uf: str = "AM",
    municipio: str = "",
    limite: int = 10,
) -> dict:
    """Busca direta no portal, sem passar por planilha."""
    uf_ok = _checar_uf(uf)
    if not uf_ok:
        return _erro("UF nao suportada: %r. Esta ferramenta cobre %s."
                     % (uf, " e ".join(UFS)))
    if uf_ok == "PB":
        return _erro(
            "a Paraiba exige coleta pelo navegador do usuario, por causa do "
            "CAPTCHA do portal; a consulta avulsa so existe para o AM"
        )
    municipio = municipio or PADRAO_MUNICIPIO[uf_ok]
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            adapter, _coord = bp.construir_adapter(
                uf_ok, municipio, 9999, 48, 2, False, 180, None)
            ofertas = adapter.buscar(termo)
    except Exception as e:
        return _erro("a consulta falhou: %s" % e, log=buf.getvalue()[-1200:])

    achados = []
    for o in sorted(ofertas, key=lambda x: x.preco)[:max(1, min(limite, 30))]:
        achados.append({
            "produto": o.descricao,
            "preco": o.preco,
            "embalagem": ("%g %s" % (o.medida.total, o.medida.base)
                          if o.medida.total else ""),
            "loja": o.estabelecimento,
            "municipio": o.municipio,
            "distancia_km": o.distancia_km,
            "data_venda": o.data_venda,
            "gtin": o.gtin,
        })
    return {
        "ok": True,
        "termo": termo,
        "municipio": municipio,
        "encontradas": len(ofertas),
        "ofertas": achados,
        "ressalva": ("Sao precos de NFC-e ja emitida, nao ofertas vigentes. "
                     "A lista nao foi filtrada por correspondencia de produto: "
                     "confira se cada item e mesmo o que voce procura."),
    }


@mcp.tool(
    title="Verificar instalacao",
    description=(
        "Roda a bateria de testes da logica (sem rede) e confere se o portal "
        "responde. Use quando algo parecer errado, antes de desconfiar dos "
        "numeros de um relatorio."
    ),
)
def verificar_instalacao(uf: str = "AM") -> dict:
    """Autoteste local + um toque no portal, para separar os dois problemas."""
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
        codigo = bp.selftest()
    saida = buf.getvalue()
    testes = {
        "passou": codigo == 0,
        "verificacoes_ok": saida.count("  ok "),
        "falhas": [l.strip() for l in saida.splitlines() if l.strip().startswith("FALHA")],
    }

    portal: dict[str, Any] = {"testado": False}
    uf_ok = _checar_uf(uf)
    if uf_ok == "AM":
        buf2 = io.StringIO()
        try:
            with contextlib.redirect_stdout(buf2), contextlib.redirect_stderr(buf2):
                adapter, _coord = bp.construir_adapter(
                "AM", "Manaus", 9999, 48, 1, False, 180, None)
                ofertas = adapter.buscar("detergente")
            portal = {"testado": True, "respondeu": True, "ofertas_no_teste": len(ofertas)}
        except Exception as e:
            portal = {"testado": True, "respondeu": False, "erro": str(e)}

    return {
        "ok": testes["passou"] and portal.get("respondeu", True),
        "testes_da_logica": testes,
        "portal": portal,
        "python": sys.version.split()[0],
        "pasta_da_skill": str(AQUI),
    }


if __name__ == "__main__":
    mcp.run(transport="stdio")
