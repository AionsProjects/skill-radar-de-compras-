# Contratos dos portais — capturados em 08/09/2026

Requisições reais, verbatim. Se o adaptador parar de funcionar, compare com o
que está aqui **antes** de reescrever qualquer coisa.

---

## AM — Busca Preço AM

```
GET  https://buscapreco.sefaz.am.gov.br/home        -> cookie JSESSIONID
POST https://buscapreco.sefaz.am.gov.br/item/grupo/p/1
     Content-Type: application/x-www-form-urlencoded
     descricaoProd=cebola&latitude=&longitude=&g-recaptcha-response=&action=
     -> HTML em latin-1
```

Rota capturada do formulário da home no Chrome. A rota antiga
`/item/grupo/page/<n>` ainda responde igual, e as duas estão em
`AdapterAM.CANDIDATOS`.

### Seletores dos cards

| Dado | Seletor |
|---|---|
| card | `div.card.small.p.hoverable` |
| nome | `.card-title b` |
| preço | `.tb-valor-25 b` |
| data (relativa) | `.tb-valor-10` (o **primeiro**; o segundo é o endereço) |
| estabelecimento | `.truncate.tooltipped.padding10 b` |
| endereço | `.valign-wrapper div span` |

Exemplo de endereço: `DUQUE DE CAXIAS, NRO 1839, PRACA 14, MANAUS-AM, CEP 69020-141`
Exemplo de data: `Há 1 dia(s) 2 hora(s) 50 minuto(s) 26 segundo(s)`

### ARMADILHA: nunca envie multipart

O servidor **não vincula** o campo em multipart. Devolve zero resultados e a
página exibe a **busca anterior guardada na sessão** — termos diferentes voltam
o mesmo produto, e o relatório atribui preços ao item errado.

Prova de que os três formatos não são equivalentes (mesma sessão, dois termos):

```
A multipart (files={'descricaoProd': (None, termo)})
   SABAO EM PO OMO  -> []          <- zero cards
   CEBOLA           -> []
B urlencoded só descricaoProd
   SABAO EM PO OMO  -> ['SABAO PO OMO 400G', ...]
   CEBOLA           -> ['BOLA DE ISOPOR 25MM', ...]
C urlencoded corpo completo (o que o navegador envia)
   SABAO EM PO OMO  -> ['SABAO PO OMO 400G', ...]
   CEBOLA           -> ['BOLA DE ISOPOR 25MM', ...]
```

O código original tinha fallback de multipart para urlencoded, mas o guard era
`"card" not in r.text` — e a palavra `card` está nas classes CSS de **qualquer**
página do portal, então o fallback nunca disparava. Por isso a checagem hoje é
`AdapterAM._tem_cards()`, que exige card de verdade no HTML parseado.

### Dados escondidos nos gatilhos do card

O card não mostra na tela, mas traz nos `onclick`:

```html
<a onclick="findByGtin(7893500018469);">                          <- codigo de barras
<a onclick="javascript:refreshMap( -59.9931023, -3.0354544);">    <- coordenada da loja
```

`refreshMap` recebe **(longitude, latitude)**, nessa ordem — confirmado por
Manaus, que fica em lat −3, lon −60. É o que permite distância exata por
estabelecimento em vez do centroide do município.

**ARMADILHA: não procure dígitos soltos para achar o GTIN.** A parte decimal da
coordenada (`-3.05139623399998`) tem 14 dígitos e passa por "EAN-14". Um
fallback assim atribuiu `05139623399998` a um feijão. Código de barras errado é
pior que nenhum: manda o usuário buscar outro produto. Só `findByGtin` vale;
sem ele, campo vazio.

### O AM não tem URL de produto

`findByGtin` apenas preenche o campo escondido `cdGtin` e submete o formulário
`#frmConsulta` por POST em `/item/grupo/page/1`. Medido em 09/09/2026:

```
GET /item/grupo/page/1?cdGtin=<codigo>   -> HTTP 500
GET /item/grupo/p/1?cdGtin=<codigo>      -> HTTP 500
GET /item/gtin/<codigo>                  -> HTTP 404
POST com descricaoProd + cdGtin          -> 200, filtra a variante exata
```

O POST funciona **sem sessão prévia**. Não prometa link de produto no AM: dê o
GTIN. (A PB, ao contrário, tem `/produtos/<slug>`.)

### Filtros que existem no formulário

Além de `descricaoProd` e `cdGtin`: `tipoConsulta` (0/24/48/168 horas),
`distancia` (2/5/10/9999 km), `municipio`, `precoMinimo`, `precoMaximo`.

O adaptador **não** usa raio e município: enviar `distancia=2` com coordenada de
Manaus não mudou o resultado, então não está provado que funcionem por POST.
Registrado como medido, não como suposto.

### Teste de não-contaminação (rode depois de qualquer mudança no adaptador)

Buscar três termos diferentes na mesma sessão tem de dar três resultados
diferentes, e repetir um termo tem de dar o mesmo resultado:

```
SABAO EM PO OMO       -> ['SABAO PO OMO 400G', ...]
CEBOLA                -> ['BOLA DE ISOPOR 25MM', ...]
LEITE ITAMBE INTEGRAL -> ['LEITE EM PO ITAMBE 200G INTEGRAL SACHET', ...]
SABAO EM PO OMO       -> ['SABAO PO OMO 400G', ...]   <- igual à primeira
```

---

## BA — Preço da Hora (aplicação antiga)

> **A PB NÃO usa mais este contrato** — foi reescrita em Next.js em 2026; veja a
> seção "PB — versão Next.js" no fim deste arquivo. O que segue vale para a
> Bahia (`precodahora.ba.gov.br`) e como registro histórico da PB.

```
GET  <base>/            -> Set-Cookie + CSRF no atributo data-id de #validate
POST <base>/sugestao/   corpo: item=<termo>
POST <base>/produtos/   corpo: gtin, latitude, longitude, raio, horas,
                               precomin=0, precomax=0, ordenar=preco.asc,
                               pagina, processo=carregar, totalRegistros=0,
                               totalPaginas=0, pageview=lista
headers: Content-Type: application/x-www-form-urlencoded; charset=UTF-8
         X-CSRFToken: <token> | Referer: <base>/
```

Base: BA `https://precodahora.ba.gov.br` (a PB migrou; ver o fim do arquivo)

O CSRF vem assim:
```html
<meta data-id="ImM0NjEzNmNjZDYzM2QxOTU4N2U2NzhjNGNhYjkxMDk3ODM0YTliOTki.aqDDww.okaneclURqbHwCMVZnzR_TKQFuo" id="validate"/>
```

### ARMADILHA: `/produtos/` é aninhado, não plano

O contrato antigo (`valor`, `nm_emp`, `nm_logr`, `nm_mun`, `distkm`) **não
existe mais**. Um adaptador que lê esses nomes devolve 0 ofertas mesmo com
HTTP 200 e 380 registros na resposta. Registro real:

```json
{"produto": {
   "codProduto": "7896098900208", "gtin": 7896098900208,
   "descricao": "DET LIQ YPE 500ML NEUTRO",
   "precoUnitario": 1.98, "precoLiquido": 1.98, "unidade": "UN",
   "data": "2026-09-09 01:55:36-00:00",
   "intervalo": "há 31 minuto(s) e 9 segundo(s)"},
 "estabelecimento": {
   "nomeEstabelecimento": "ATAKADAO ATAKAREJO",
   "endLogradouro": "AVENIDA SANTOS DUMONT", "endNumero": "5840",
   "bairro": "PITANGUEIRAS", "municipio": "LAURO DE FREITAS", "uf": "BA",
   "cnpj": "73849952000310",
   "latitude": -12.8767219, "longitude": -38.308648, "distancia": 22.0842}}
```

`AdapterPrecoDaHora._ler_item` aceita as duas formas, aninhada primeiro.

### `horas` é em horas; `dias` na resposta é só o eco

Medido:

| enviado | eco `dias` | totalRegistros | janela observada |
|---|---|---|---|
| `horas=72` | 72 | 380 | ~14 h na página 1 (25 mais baratos) |
| `horas=3` | 3 | 57 | 23:42 → 01:55 |
| `dias=1` | 72 | 380 | ignorado (voltou ao padrão) |

Ou seja: `horas=72` = 3 dias, como documentado. O nome `dias` na resposta é
enganoso. Enviar `dias` no corpo não tem efeito.

`raio` funciona: `raio=5` → 122 registros, maior distância 4,67 km.
`ordenar=preco.asc` funciona.

### Encoding: não "corrija"

As descrições vêm em UTF-8 correto. `DETERGENTE L\xc3\x8dQUIDO` é `Í`
(LATIN CAPITAL LETTER I WITH ACUTE). Se aparecer `L?QUIDO` no terminal, é o
console do Windows em cp1252 — o dado está íntegro. Não troque o decode.

### PB: Cloudflare com CAPTCHA

Em 08/09/2026, HTTP puro recebe 403 e o Chrome real é redirecionado para
`https://precodahora.tcepb.tc.br/cf_captcha?urlBack=/`, com título "Um momento…"
e sem `#validate` na página. Não é bloqueio por IP: é desafio interativo.

O único caminho é `--navegador`, com **uma pessoa** resolvendo o desafio. A BA
não tem anti-bot e serve para validar o fluxo do código quando a PB está
inacessível:

```
python comparador_preco_sefaz.py --smoke --uf BA --municipio Salvador --termo detergente
```

BA funcionando + PB em 403 = código certo, host bloqueado.

---

## Como recapturar uma requisição, se algo mudar

O AM não tem anti-bot, então dá para instrumentar com Playwright:

```python
from playwright.sync_api import sync_playwright
capturadas = []
with sync_playwright() as pw:
    nav = pw.chromium.launch(channel="chrome", headless=False)
    pag = nav.new_page()
    pag.on("request", lambda r: capturadas.append(
        {"metodo": r.method, "url": r.url,
         "content_type": r.headers.get("content-type", ""),
         "post_data": (r.post_data or "")[:500]}) if r.method == "POST" else None)
    pag.goto("https://buscapreco.sefaz.am.gov.br/home", wait_until="networkidle")
    capturadas.clear()                     # só o que vier da busca interessa
    campo = pag.query_selector("#descricaoProd")
    campo.fill("cebola"); campo.press("Enter")
    pag.wait_for_timeout(7000)
    print(capturadas)                      # método, caminho, campo, content-type
    nav.close()
```

Anote o resultado **verbatim** neste arquivo, com a data.

---

## PB — Preço da Hora Paraíba, versão Next.js (09/09/2026)

**A PB foi reescrita.** O contrato desta seção substitui o da seção anterior
para `precodahora.tcepb.tc.br`. A Bahia continua com a aplicação antiga.

Como reconhecer a versão nova em 5 segundos:

```js
document.head.innerHTML.includes('/_next/')   // true = Next.js, contrato novo
document.querySelector('#validate')           // null  = contrato antigo morreu
```

### As duas rotas

Os hashes **mudam a cada deploy** — nunca fixe. Capture com hook em
`window.fetch` (procedimento no SKILL.md, seção 3 da PB).

```
POST /api/<hashSugestao>
     content-type: application/json
     request-hac: 70cd06da2552134335551acca82fd4c0
     request-id:  c05de047aa0574962c8bc195eb686fd5
     {"content":"detergente ype"}

  -> {"success":true,"data":[
       {"name":"DETERGENTE YPE LIQ 500ML NEUTRO",
        "slug":"detergente-ype-liq-500ml-neutro",
        "id":7896098900208, "score":16.424}, ...]}
```

`id` é o **GTIN**. `slug` compõe a URL do produto — a PB **tem** deep link:
`https://precodahora.tcepb.tc.br/produtos/detergente-ype-liq-500ml-neutro`

```
POST /api/<hashPrecos>
     mesmos headers
     {"id":7896098900208,"distance":10,"page":1,"sort":"lowestPrice",
      "lat":-7.1194958,"lng":-34.8450118,"regionId":1}

  -> {"success":true,"data":{
       "pagination":{"currentPage":1,"totalItems":168,"totalPages":4,"pageItems":50},
       "days":[{"total":45,"id":"2026-09-09","name":"2026-09-09"}, ...],
       "priceRange":{"min":1.99,"max":4.49,"avg":2.64},
       "cities":[{"name":"JOAO PESSOA","total":152,"id":"2507507"},
                 {"name":"BAYEUX","total":11,"id":"2501807"}, ...],
       ...}}
```

### ARMADILHA: request-hac e request-id são obrigatórios

Medido, na mesma sessão e com o mesmo corpo:

```
sem os dois headers   -> HTTP 401 {"success":false,"message":"Unauthorized"}
reusando os dois      -> HTTP 200 (funcionam para chamadas seguintes)
```

São gerados pelo JavaScript do app. **Não reproduza essa geração** — é proteção
anti-abuso. Colete de dentro da página, onde o app os emite. Foi por isso, e
não só pelo Cloudflare, que nenhum cliente HTTP puro ia funcionar aqui.

Observado: a validade é curta e chamadas repetidas fora do fluxo do app às vezes
voltam sem `data`. Capture headers frescos se começar a falhar.

### Extrator de DOM — o caminho que sempre funciona

A página do produto lista as lojas com nome, preço, data e bairro/cidade. As
classes CSS são hasheadas (`c-knNdqq`), então **não** se ancore nelas: identifique
o card pelo conteúdo. Foi assim que a primeira coleta real da PB saiu.

Navegue para `/produtos/<slug>`, **role até o fim** (a lista carrega ao rolar) e
rode:

```js
window.__extrair = function () {
  const RE_DATA = /(\d{2}\/\d{2}\/\d{4}),?\s*(\d{2}:\d{2})?/;
  const todos = Array.from(document.querySelectorAll('div,li,article'));
  const cands = todos.filter(e => {
    const t = e.innerText || '';
    return t.includes('Preço') && t.includes('R$') && RE_DATA.test(t) && t.length < 300;
  });
  const folhas = cands.filter(c => !cands.some(o => o !== c && c.contains(o)));
  const gtinEl = document.body.innerText.match(/\b(\d{13})\b/);
  const titulo = (document.querySelector('h1') || {}).innerText
    || (document.title.split('|').pop() || '').trim();
  const lojas = folhas.map(f => {
    const linhas = (f.innerText || '').split('\n').map(s => s.trim()).filter(Boolean);
    const iPreco = linhas.indexOf('Preço');
    let preco = null;
    if (iPreco >= 0) {
      const nums = linhas.slice(iPreco + 1, iPreco + 4).filter(s => /^\d+$/.test(s));
      if (nums.length >= 2) preco = parseFloat(nums[0] + '.' + nums[1]);
      else {
        const m = (f.innerText || '').match(/R\$\s*([\d.,]+)/);
        if (m) preco = parseFloat(m[1].replace('.', '').replace(',', '.'));
      }
    }
    const mData = (f.innerText || '').match(RE_DATA);
    const local = linhas[linhas.length - 1] || '';
    const partes = local.split(',').map(s => s.trim());
    return {
      loja: linhas[0] || '', preco: preco,
      data: mData ? (mData[1] + (mData[2] ? ' ' + mData[2] : '')) : '',
      bairro: partes.length > 1 ? partes.slice(0, -1).join(', ') : '',
      cidade: partes[partes.length - 1] || '',
    };
  }).filter(l => l.preco !== null && l.loja);
  return {produto: titulo, gtin: gtinEl ? gtinEl[1] : '', url: location.href, lojas};
};
window.__extrair()
```

O preço vem em linhas separadas no DOM — `"Preço"`, `"R$"`, `"1"`, `"99"` — daí
a montagem por índice em vez de um regex sobre o texto todo.

Acumule entre navegações em `localStorage` (`__coleta`), porque cada produto é
uma navegação. Guarde também `__extrair.toString()` e reinjete com `eval` se a
função sumir.

### Diferenças materiais que o relatório precisa declarar

| | AM | PB (Next.js) |
|---|---|---|
| Natureza do preço | valor de **uma NFC-e**, com hora | **média das últimas vendas** do lojista |
| Janela | 48 h | **40 dias** ("Baseado nos últimos 40 dias") |
| Link de produto | não existe | `/produtos/<slug>` |
| Coordenada da loja | no card, via `refreshMap` | não exposta no DOM |
| Paginação | 12 cards por página | 50 por página (`totalPages`) |

A página diz textualmente: *"Os preços são médias baseadas nas últimas vendas de
cada produto pelos lojistas"*. O módulo marca a fonte como
`Preco da Hora PB (media do lojista)` — não apresente como preço de nota.

### Cloudflare

No **Chrome normal do usuário** (via extensão), a verificação passa sozinha. Em
Chrome sob automação (Playwright), aparece o checkbox "Confirme que é humano" do
Turnstile, dentro de shadow DOM — `input[type=checkbox]` não o encontra. Quem
clica é o usuário; com `--perfil`, o cookie fica guardado para as próximas.

Quando a pessoa clica, o Cloudflare **navega** da `/cf_captcha` para a home, e um
`page.evaluate` em curso morre com *"Execution context was destroyed"*. Isso é o
**sucesso** da verificação, não a falha dela — trate a exceção e tente de novo
no ciclo seguinte.
