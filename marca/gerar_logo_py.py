# -*- coding: utf-8 -*-
"""
Gera marca/logo_aions.py a partir dos paths da logo, para o reportlab desenhar.

NAO redesenha a marca: os paths vem do arquivo de marca (logo-aions.svg), pelo
mesmo caminho ja usado no relatorio da Esteira Financeira. Para trocar a logo,
troque o SVG e rode este gerador -- nunca edite o .py na mao.
"""
import io
import json
import re

dados = json.load(io.open("_logo_partes.json", encoding="utf-8"))

linhas = [
    '# -*- coding: utf-8 -*-',
    '"""',
    'Logo da AIONS em vetor, para desenhar no PDF com reportlab.',
    '',
    'GERADO por marca/gerar_logo_py.py a partir de marca/logo-aions.svg.',
    'Nao editar a mao e nao redesenhar a marca: troque o SVG e rode o gerador.',
    'Os paths usam so os operadores m/l/c/h, que mapeiam 1:1 para o reportlab.',
    '"""',
    '',
    'LARGURA = %r' % dados["w"],
    'ALTURA = %r' % dados["h"],
    '',
    '# (cor RGB 0-1, preenchimento even-odd?, path)',
    'PARTES = [',
]
for p in dados["partes"]:
    linhas.append("    (%r, %r," % (tuple(round(c, 4) for c in p["cor"]), p["eo"]))
    d = p["d"]
    # quebra o path em pedacos legiveis, sem alterar o conteudo
    for i in range(0, len(d), 88):
        linhas.append("     %r" % d[i:i + 88] + ("," if i + 88 >= len(d) else ""))
    linhas.append("     ),")
linhas.append("]")
linhas += [
    '',
    '',
    'def desenhar(canv, x, y_topo, altura):',
    '    """',
    '    Desenha a logo com o topo em `y_topo` (medido do TOPO da pagina, como no',
    '    resto do relatorio) e devolve a largura ocupada.',
    '',
    '    O eixo Y do PDF cresce para cima e o do SVG para baixo, por isso a escala',
    '    em Y e negativa -- e o mesmo ajuste do gerador em JS.',
    '    """',
    '    from reportlab.pdfgen.canvas import FILL_EVEN_ODD, FILL_NON_ZERO',
    '',
    '    escala = altura / ALTURA',
    '    pagina_altura = canv._pagesize[1]',
    '    canv.saveState()',
    '    canv.translate(x, pagina_altura - y_topo)',
    '    canv.scale(escala, -escala)',
    '    for cor, even_odd, d in PARTES:',
    '        canv.setFillColorRGB(*cor)',
    '        p = canv.beginPath()',
    '        _traçar(p, d)',
    '        canv.drawPath(p, fill=1, stroke=0,',
    '                      fillMode=FILL_EVEN_ODD if even_odd else FILL_NON_ZERO)',
    '    canv.restoreState()',
    '    return altura * LARGURA / ALTURA',
    '',
    '',
    'def _traçar(p, d):',
    '    """Converte os operadores de path do PDF em chamadas do reportlab."""',
    '    fichas = d.replace("\\n", " ").split()',
    '    pilha = []',
    '    for f in fichas:',
    '        if f == "m":',
    '            p.moveTo(float(pilha[-2]), float(pilha[-1])); pilha = []',
    '        elif f == "l":',
    '            p.lineTo(float(pilha[-2]), float(pilha[-1])); pilha = []',
    '        elif f == "c":',
    '            v = [float(x) for x in pilha[-6:]]',
    '            p.curveTo(v[0], v[1], v[2], v[3], v[4], v[5]); pilha = []',
    '        elif f == "h":',
    '            p.close(); pilha = []',
    '        else:',
    '            pilha.append(f)',
    '',
]
io.open("logo_aions.py", "w", encoding="utf-8").write("\n".join(linhas) + "\n")
print("logo_aions.py gerado: %d partes" % len(dados["partes"]))
