---
name: busca-preco
description: Compara os preços de uma planilha de produtos com os preços realmente praticados no comércio, pelos portais oficiais de preço por NFC-e do Amazonas (Busca Preço AM) e da Paraíba (Preço da Hora PB). Usar quando o usuário invocar /busca-preco; pedir para comparar preços de uma planilha, cotação ou lista de compras contra o mercado; conferir se um fornecedor está cobrando caro; pesquisar preço de produto no Amazonas ou na Paraíba; ou mencionar NFC-e, SEFAZ AM, Busca Preço, Preço da Hora, precodahora. Gera planilha anotada .xlsx e PDF explicativo. No AM consulta direto; na PB usa a extensão do Chrome do usuário, que resolve o CAPTCHA. Não emite veredito de "vale a pena": reporta economia e distância lado a lado.
---

# busca-preco

Preço de nota fiscal, não de anúncio. Cada número vem de uma NFC-e efetivamente
emitida por um estabelecimento e enviada à SEFAZ — é preço que alguém
**realmente pagou** no passado recente.

O script fica nesta pasta. Chame por caminho absoluto e deixe as saídas no
diretório do usuário:

```bash
BP="$HOME/.claude/skills/busca-preco/busca_preco.py"
```

## Como ler os argumentos

O usuário invoca `/busca-preco <argumentos>` em linguagem livre. Extraia:

| O que | Como reconhecer | Se faltar |
|---|---|---|
| **planilha** | caminho de `.xlsx`, `.csv` ou `.tsv`, ou anexo enviado | é obrigatório: peça o arquivo |
| **uf** | `AM`/`Amazonas`, `PB`/`Paraíba`; ou o município citado | pergunte, com AskUserQuestion |
| **município** | nome de cidade | AM→Manaus, PB→João Pessoa; diga qual assumiu |

Exemplos que devem funcionar sem mais perguntas:

- `/busca-preco lista.xlsx AM` → Amazonas, referência Manaus
- `/busca-preco compras.xlsx Manaus` → UF vem do município
- `/busca-preco cotacao.csv Paraíba João Pessoa` → caminho da PB, com CAPTCHA

Se o usuário só passar a planilha, **não invente a UF**: esta skill cobre dois
estados, e a resposta muda inteiramente entre eles. Pergunte.

## Antes de qualquer consulta

```bash
pip install requests beautifulsoup4 lxml openpyxl reportlab
python "$BP" --selftest
```

64 verificações de lógica pura, sem rede, várias delas regressões de defeitos
reais. **Se alguma falhar, pare e corrija antes de consultar os portais** — um
teste vermelho aqui significa que os números do relatório não valem nada.

## Amazonas — direto, sem navegador

```bash
python "$BP" --smoke --uf AM --termo detergente          # o portal responde?
python "$BP" --uf AM --municipio Manaus \
    --planilha lista.xlsx --saida cotacao --brutos brutos/
```

Sem login, sem captcha. Janela de 48 h.

O formulário tem filtros de `tipoConsulta` (0/24/48/168 h), `distancia`
(2/5/10/9999 km), `municipio`, `precoMinimo` e `precoMaximo`. O adaptador **não
usa** os de raio e município: enviar `distancia=2` com coordenada de Manaus não
mudou o resultado (medido em 09/09/2026), então não afirme que funcionam por
POST. O município de cada oferta sai do endereço do estabelecimento.

Dois dados ficam escondidos nos gatilhos de JavaScript do card, e o adaptador os
lê: `findByGtin(<codigo>)` dá o **código de barras**, e
`refreshMap(<lon>, <lat>)` dá a **coordenada da loja** — que é o que permite
distância exata por estabelecimento, em vez do centroide do município. A data
vem relativa ("Há 7 hora(s)") e é convertida para absoluta no fuso de Manaus.

**O AM não tem URL de produto.** `findByGtin` só preenche um campo e submete um
formulário por POST; `GET ...?cdGtin=` devolve 500 e `/item/gtin/<codigo>` dá
404 (medido em 09/09/2026). Não prometa link de produto no AM — dê o GTIN.

Use `--brutos DIR` sempre que o resultado for embasar uma decisão: é o que
permite rastrear qualquer número até a resposta do portal.

## Paraíba — pela extensão do Chrome do usuário

**O portal da PB foi reescrito em Next.js** (verificado em 09/09/2026). O
contrato antigo — `#validate`, `POST /sugestao/`, `POST /produtos/` em
form-urlencoded — **não existe mais**. Isso vale só para a PB: a Bahia
(`precodahora.ba.gov.br`) segue com a aplicação antiga.

A API nova exige os headers `request-hac` e `request-id`, gerados pelo
JavaScript do próprio app. **Sem eles a resposta é 401.** Não tente reproduzir
essa geração: é proteção anti-abuso. Colete de dentro da página, onde o app os
produz — é para isso que a extensão serve.

O Cloudflare, no Chrome normal do usuário, passa sozinho. Se aparecer o
checkbox "Confirme que é humano", **quem clica é o usuário**; você não resolve
verificação anti-bot.

### 1. Descubra o que consultar (não toca na rede)

```bash
python "$BP" --uf PB --municipio "Joao Pessoa" \
    --planilha lista.xlsx --listar-termos > termos.json
```

Use `termos_unicos`. A referência de localização do portal é por **CEP**, e
aparece no topo da página (ex. `58010000 VARADOURO, JOÃO PESSOA`); confira se
corresponde ao município pedido antes de coletar.

### 2. Abra o portal

```
ToolSearch "select:mcp__claude-in-chrome__tabs_context_mcp,mcp__claude-in-chrome__tabs_create_mcp,mcp__claude-in-chrome__navigate,mcp__claude-in-chrome__javascript_tool,mcp__claude-in-chrome__computer,mcp__claude-in-chrome__find,mcp__claude-in-chrome__browser_batch"
```

`tabs_context_mcp` → `tabs_create_mcp` → `navigate` para
`https://precodahora.tcepb.tc.br/`. Confirme que passou:

```js
({url: location.href, titulo: document.title})
```

Título com "Preço da Hora" e URL sem `cf_captcha` = liberado. Se estiver em
`cf_captcha`, peça o clique ao usuário em uma frase e espere a resposta dele.

**Não clique no banner de cookies.** Ele só oferece "Aceitar", e aceitar termos
em nome do usuário exige autorização dele. A busca funciona sem isso; o banner
pode interceptar cliques por coordenada, então prefira `find` + clique por
`ref`, ou `form_input`.

### 3. Descubra as rotas desta build

Os hashes das rotas **mudam a cada deploy**, então não os fixe: capture-os. Um
hook em `window.fetch` antes de interagir resolve:

```js
window.__cap = [];
if (!window.__of) window.__of = window.fetch;
window.fetch = async function (...a) {
  const url = a[0] instanceof Request ? a[0].url : String(a[0]);
  const r = await window.__of.apply(this, a);
  if (url.includes('/api/')) {
    let h = {};
    try { new Headers((a[1] || {}).headers || {}).forEach((v, k) => h[k] = v); } catch (e) {}
    window.__cap.push({url: url.replace(location.origin, ''), hdrs: h,
                       corpo: (a[1] || {}).body || ''});
  }
  return r;
};
'hook instalado'
```

Agora digite um termo no campo de busca (`find` "campo Digite sua busca" →
clique por `ref` → `type`) e espere 5 s: isso dispara a **rota de sugestão**.
Clique na primeira sugestão e espere ~9 s: isso dispara a **rota de preços**.
Depois leia:

```js
window.__cap.map(c => ({url: c.url, corpo: String(c.corpo).slice(0, 120), hdrs: c.hdrs}))
```

Você vai ver duas rotas `/api/<hash>` e os headers a reusar:

```
POST /api/<hashSugestao>   {"content":"detergente ype"}
  → {"success":true,"data":[{"name":"DETERGENTE YPE LIQ 500ML NEUTRO",
       "slug":"detergente-ype-liq-500ml-neutro","id":7896098900208,"score":16.4}]}
     -- `id` é o GTIN e `slug` dá a URL do produto: /produtos/<slug>

POST /api/<hashPrecos>     {"id":7896098900208,"distance":10,"page":1,
                            "sort":"lowestPrice","lat":-7.1194958,
                            "lng":-34.8450118,"regionId":1}
  → {"success":true,"data":{"pagination":{"totalItems":168,"totalPages":4,
       "pageItems":50}, "priceRange":{"min":1.99,"max":4.49,"avg":2.64},
       "cities":[...], "days":[...], ...}}

headers dos dois: content-type: application/json
                  request-hac: <32 hex>   request-id: <32 hex>
```

Os headers capturados são **reutilizáveis** para as chamadas seguintes.

### 4. Colete, uma consulta por vez

Com as duas rotas e os headers, itere os termos de `termos.json` em **um**
`javascript_tool`, com pausa de 1,5 s. É infraestrutura pública — não martele.

```js
const S = window.__cap.find(c => String(c.corpo).includes('content'));
const P = window.__cap.find(c => String(c.corpo).includes('"id"'));
const TERMOS = ["DETERGENTE YPE NEUTRO", "ARROZ TIO JOAO"];   // de termos.json
const {lat, lng, regionId} = JSON.parse(P.corpo);
const espera = ms => new Promise(r => setTimeout(r, ms));
const post = async (rota, hdrs, obj) => {
  const r = await fetch(rota, {method: 'POST', headers: hdrs,
                               body: JSON.stringify(obj), credentials: 'include'});
  return r.ok ? r.json() : {erro: r.status};
};
const saida = {};
for (const termo of TERMOS) {
  const sug = await post(S.url, S.hdrs, {content: termo});
  const cand = ((sug.data) || [])[0];
  if (!cand) { saida[termo] = {produto: '', gtin: '', lojas: []}; await espera(1500); continue; }
  await espera(1500);
  const pr = await post(P.url, P.hdrs, {id: cand.id, distance: 10, page: 1,
                          sort: 'lowestPrice', lat, lng, regionId});
  const d = pr.data || {};
  const lista = d.items || d.products || d.establishments || d.results || [];
  saida[termo] = {
    produto: cand.name, gtin: String(cand.id),
    url: location.origin + '/produtos/' + cand.slug,
    lojas_vistas: (d.pagination || {}).totalItems || lista.length,
    lojas: lista.map(x => [
      x.name || x.establishment || (x.store || {}).name || '',
      x.price ?? x.value ?? x.lowestPrice,
      x.date || x.saleDate || '',
      x.district || x.neighborhood || '',
      x.city || (x.address || {}).city || ''])
  };
  await espera(1500);
}
JSON.stringify(saida);
```

**Confira a forma da lista antes de confiar no `map`**: os nomes de campo acima
são tentativas. Rode primeiro `Object.keys(d)` e inspecione um item; ajuste, e
anote o que encontrou em `referencias/contratos-portais.md`.

**Alternativa que sempre funciona**: extrair do DOM da página do produto. A
página lista as lojas com nome, preço, data e bairro/cidade, e foi assim que a
primeira coleta real saiu. Navegue para `/produtos/<slug>`, role até o fim para
carregar a lista, e leia os cards — o extrator está em
`referencias/contratos-portais.md`.

Grave o retorno em `coletado.json` (ferramenta Write) **sem editar**.

### 5. Processe com o mesmo motor do AM

```bash
python "$BP" --uf PB --municipio "Joao Pessoa" \
    --planilha lista.xlsx --ofertas-json coletado.json --saida cotacao
```

O formato aceito é `{termo: {produto, gtin, url, lojas: [[loja, preco, data,
bairro, cidade], ...]}}` — cada loja como lista ou como objeto
`{loja, preco, data, bairro, cidade}`. Chaves começando com `_` são tratadas
como metadado. O contrato antigo (`{termo: [resposta de /produtos/]}`) continua
aceito, para a Bahia.

### O que é diferente da PB para o AM — e precisa ser dito no relatório

| | AM | PB |
|---|---|---|
| Natureza do preço | valor de **uma NFC-e**, com hora | **média das últimas vendas** do lojista |
| Janela | 48 h | **40 dias** |
| Link de produto | não existe | `/produtos/<slug>` |
| Coordenada da loja | vem no card (distância exata) | não vem no DOM; distância cai no centroide do município |

O portal da PB diz, na própria página: *"Os preços são médias baseadas nas
últimas vendas de cada produto pelos lojistas"*. **Não apresente isso como
preço de nota individual.** O módulo marca a fonte como
`Preco da Hora PB (media do lojista)`; preserve essa distinção ao relatar.

Raio padrão 10 km (`distance`). Não use `precodahora.pb.gov.br`, fora do ar
desde 2023.

## Regras de conduta

- Uma consulta por vez, pausa de 1–2 s, backoff de 30 s em 429/503. É
  infraestrutura pública de SEFAZ — não martele.
- **Não contorne a verificação anti-bot.** Sem solvers, sem stealth patches, sem
  fingerprint forjado. O usuário resolve, ou não há PB.
- Se um número parecer bom demais, **desconfie do casamento de descrição antes
  de comemorar**. Foi assim que todos os defeitos abaixo apareceram.
- Ao relatar, diga o que não foi exercitado. Cobertura de teste não é cobertura
  de realidade.

## O que a skill se recusa a fazer

Cada regra existe porque já produziu um número errado num relatório real. Não
relaxe nenhuma para "melhorar" o resultado.

1. **Comparar preço absoluto entre embalagens diferentes.** Tudo por unidade
   base (R$/litro, R$/quilo). 1 L a R$ 1,50 e 500 ml a R$ 0,75 custam o mesmo, e
   a skill reporta economia zero.
2. **Comparar bases diferentes.** 1 L de leite líquido não se compara com 400 g
   de leite **em pó**: economia não calculada, com a razão na linha.
3. **Aceitar casamento por trecho de texto.** Os portais casam por aproximação:
   `CEBOLA` puxa `COLA` e `BOLA`, devolvendo cola de silicone, bola de isopor,
   Coca-Cola e `SALGADINHO … CEBOLA E SALSA`. Resultado sem palavra em comum, ou
   cuja descrição é majoritariamente outro produto (Dice < 0,35), é descartado
   antes de qualquer conta.
4. **Tomar preço fora da curva como preço de mercado.** O portal do AM devolveu
   NFC-e real de lata de Coca-Cola a **R$ 0,01** num restaurante — brinde ou
   ajuste fiscal. Sozinha, virava "99,43% de economia". Oferta abaixo de 15% da
   mediana das outras do mesmo produto é descartada e declarada. Não é filtro de
   preço baixo: metade da mediana passa, porque economia real existe.
5. **Calcular economia em casamento BAIXA.** O preço aparece como referência
   sinalizada; a economia fica em branco para conferência humana.
6. **Inventar preço.** Produto ausente vira `NAO_ENCONTRADO` — nunca preço
   estimado, nunca preço de e-commerce.
7. **Inventar distância.** Linha reta por haversine; geocodificação que falha
   vira `n/d`. No AM, oferta de outro município leva ressalva de acesso fluvial:
   300 km em linha reta podem ser mais de um dia de viagem de barco.
8. **Emitir veredito.** A skill não recomenda trocar de fornecedor ou de
   município. Mostra economia e distância lado a lado; a decisão é do usuário.

Detalhe conhecido que **não** é defeito: o menor preço por litro pode ser um
fardo (um "24 X 500ML." de R$ 58,00). O cálculo está certo, e a coluna
`Produto encontrado no portal` mostra a embalagem real — não esconda isso
resumindo só o valor.

## Saídas

`<saida>.xlsx` (abas Comparativo, Resumo, Auditoria), `<saida>.pdf` (A4) e
`<saida>.json` com os dados crus. Entregue os arquivos ao usuário com
SendUserFile, não só o caminho.

## Ao relatar

1. Autoteste rodado, com a contagem.
2. Quantos itens tiveram oferta, quantos ficaram BAIXA, quantos
   `NAO_ENCONTRADO` — nomeando estes.
3. Quantos resultados foram descartados: por não corresponderem ao produto e por
   preço fora da curva.
4. A economia somada, deixando claro que é **por unidade** e que não inclui
   frete nem diferença de ICMS.
5. O que ficou de fora, com todas as letras.

## Contratos dos portais

Em `referencias/contratos-portais.md`: requisições reais verbatim (capturadas em
08/09/2026), a armadilha do multipart no AM, o teste de não-contaminação e como
recapturar com navegador se uma rota mudar. Planilha de teste com as armadilhas
conhecidas em `referencias/planilha-exemplo.xlsx`.
