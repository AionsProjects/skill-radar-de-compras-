# busca-preco — Radar de Compras

Compara os preços de uma planilha de produtos com os preços **realmente
praticados** no comércio, usando os portais oficiais de preço por NFC-e do
**Amazonas** e da **Paraíba**. Devolve planilha anotada (`.xlsx`) e PDF
explicativo, com fornecedor, endereço, distância e código de barras de cada
opção.

Cada número vem de nota fiscal eletrônica emitida por um estabelecimento e
enviada à SEFAZ. Não é preço de anúncio, e a ferramenta **não recomenda** trocar
de fornecedor: mostra economia e distância lado a lado, e a decisão é sua.

## Instalar

Precisa de **Python 3.10+** e do **Claude Code**.

### Windows (PowerShell)

```powershell
git clone https://github.com/AionsProjects/skill-radar-de-compras-.git "$env:USERPROFILE\.claude\skills\busca-preco"
pip install requests beautifulsoup4 lxml openpyxl reportlab
python "$env:USERPROFILE\.claude\skills\busca-preco\busca_preco.py" --doctor
```

### macOS / Linux

```bash
git clone https://github.com/AionsProjects/skill-radar-de-compras-.git ~/.claude/skills/busca-preco
pip install requests beautifulsoup4 lxml openpyxl reportlab
python ~/.claude/skills/busca-preco/busca_preco.py --doctor
```

**O nome da pasta tem de ser `busca-preco`** — é ele que define o comando
`/busca-preco`, não o conteúdo do arquivo.

Depois de clonar, **reinicie o Claude Code** para a skill aparecer.

### Conferir se ficou tudo certo

```
python busca_preco.py --doctor
```

Verifica versão do Python, as cinco dependências, acesso ao portal do Amazonas
e roda as 74 verificações de lógica. Se algo faltar, ele diz exatamente qual
comando resolve. Termine só quando aparecer **"Tudo pronto"**.

## Usar

No Claude Code:

```
/busca-preco minha-lista.xlsx AM
```

Funciona em linguagem livre — `/busca-preco compras.xlsx Manaus`,
`/busca-preco cotacao.csv Paraíba`. Se você não disser o estado, ele pergunta,
porque a resposta muda inteiramente entre AM e PB.

### A planilha

Só precisa de duas colunas: uma de **produto** e uma de **preço**. Os nomes são
detectados automaticamente (`Produto`, `Descrição`, `Item`… / `Preço`, `Valor`,
`Custo`…), e colunas extras são ignoradas.

| Produto | Preço Atual | Fornecedor Atual |
|---|---|---|
| ARROZ TIO JOAO TIPO 1 1KG | R$ 8,90 | Distribuidora Rio Negro |
| DETERGENTE YPE NEUTRO 500ML | R$ 3,49 | Higiluz Comercial |

Aceita `.xlsx`, `.csv` e `.tsv`. Há um exemplo pronto em
`referencias/exemplo-amazonas.xlsx`, com 14 itens reais de Manaus.

### Fora do Claude Code, direto no terminal

```bash
python busca_preco.py --uf AM --municipio Manaus \
    --planilha lista.xlsx --saida cotacao --brutos brutos/
```

Sai `cotacao.xlsx` (abas Comparativo, Resumo, Alternativas, Auditoria),
`cotacao.pdf` e `cotacao.json`. `--brutos` guarda a resposta crua de cada
consulta, para rastrear qualquer número até o que o portal respondeu.

### Como ler as colunas de alternativa

A aba **Comparativo** traz, na linha de cada item, o menor preço e mais três
opções: `Alternativa 1 (2ª melhor)`, `Alternativa 2 (3ª melhor)` e
`Alternativa 3 (4ª melhor)`. A alternativa 1 é a **segunda** melhor, porque a
primeira já são as colunas de menor preço — não é a mesma loja repetida.

Cada alternativa tem duas colunas de valor, e a ordem delas é intencional:

- **"equivale a"** — quanto custaria a embalagem da **sua** planilha a esse
  preço por litro/quilo. É por esta coluna que as opções estão ordenadas, e é
  ela que se compara com o seu preço atual.
- **"preço da embalagem"** — o que a loja realmente cobra pelo que ela vende.

Os dois valores diferem quando a embalagem é outra, e aí a coluna de etiqueta
**não** fica em ordem crescente. Não é erro: um fardo de detergente com 24
unidades custa R$ 58,00 e ainda assim é o mais barato por litro (R$ 2,42 o
equivalente a 500 ml). Para saber o que você vai pagar no caixa, olhe o preço
da embalagem; para decidir se vale, olhe o equivalente.

A aba **Alternativas** traz até 5 opções por item, uma linha cada, com endereço
completo, data da venda, código de barras e link do mapa.

## Os dois estados são diferentes

|  | Amazonas | Paraíba |
|---|---|---|
| Como consulta | direto, sem navegador | **exige a extensão do Claude no Chrome** |
| Natureza do preço | valor de **uma nota**, com hora | **média das últimas vendas** do lojista |
| Janela | 48 horas | 40 dias |

**O Amazonas funciona logo depois de instalar.**

A **Paraíba** fica atrás de proteção anti-bot e de uma API que exige headers
gerados pelo próprio site. Por isso a coleta é feita pelo seu Chrome, com a
[extensão do Claude](https://claude.ai/chrome) instalada e logada na mesma
conta. Se aparecer um "Confirme que é humano", **você** clica — o Claude não
resolve verificação anti-bot. O passo a passo está no `SKILL.md`.

## O que a ferramenta se recusa a fazer

Cada uma destas regras existe porque já produziu um número errado num relatório
real. Elas são o motivo de confiar na saída.

1. **Não compara preço absoluto entre embalagens diferentes.** Tudo em R$/litro
   ou R$/quilo. 1 L a R$ 1,50 e 500 ml a R$ 0,75 custam o mesmo, e a economia
   reportada é zero.
2. **Não compara litro com quilo.** Leite líquido 1 L contra leite em pó 400 g
   não gera economia — gera uma ressalva.
3. **Não aceita casamento por trecho de texto.** Os portais casam por
   aproximação: "CEBOLA" puxa "COLA" e "BOLA", e devolve cola de silicone e bola
   de isopor. Isso é descartado antes de qualquer conta.
4. **Não toma preço fora da curva como preço de mercado.** O portal devolveu uma
   nota real de lata de refrigerante a **R$ 0,01** (brinde ou ajuste fiscal).
   Sozinha, ela virava "99% de economia".
5. **Não inventa preço.** Produto ausente vira `NAO_ENCONTRADO`, nunca
   estimativa.
6. **Não inventa distância.** Linha reta entre coordenadas; sem coordenada, fica
   "n/d". No Amazonas, oferta de outro município leva ressalva de acesso
   fluvial — 300 km em linha reta podem ser um dia de barco.
7. **Não emite veredito.** Mostra economia e distância; a decisão é de quem lê.

Uma coisa que **não** é defeito: o menor preço por litro pode ser um fardo (um
"24 X 500ML." de R$ 58,00). A conta está certa, e a coluna *Produto encontrado
no portal* mostra a embalagem real. Confira antes de comemorar.

## Antes de decidir com esses números

- Preço de nota é do **passado recente**; o estabelecimento não é obrigado a
  mantê-lo, e o produto pode estar sem estoque.
- A economia é **por unidade**. Multiplique pelas quantidades que você compra.
- **Não inclui frete nem diferença de ICMS.** Cinco fornecedores mais baratos
  podem significar cinco compras.
- Se um número parecer bom demais, olhe a coluna do produto encontrado antes de
  acreditar.

## Cobertura

Amazonas e Paraíba. A Bahia funciona como bônus (`--uf BA`, mesmo código-base da
PB antiga) e serve para diagnosticar problemas. Paraná e Alagoas têm API própria
e mais simples; os outros estados só têm o app Menor Preço Brasil, sem web nem
API pública.

## Arquivos

```
SKILL.md                          o que o Claude lê: fluxo dos dois estados
busca_preco.py                    o módulo (--doctor, --selftest, --smoke)
referencias/contratos-portais.md  as requisições reais dos portais, verbatim
referencias/exemplo-amazonas.xlsx planilha de exemplo, 14 itens de Manaus
referencias/planilha-exemplo.xlsx planilha curta com as armadilhas conhecidas
```

Se um portal mudar e a skill parar de funcionar, `contratos-portais.md` tem as
requisições reais e o procedimento para recapturá-las com o navegador.

## Manutenção

Os portais mudam sem aviso: a Paraíba foi reescrita em Next.js e todo o contrato
antigo morreu. Ao consertar, **rode `--selftest` antes e depois** e registre em
`referencias/contratos-portais.md` o que você mediu, com data. As 74
verificações não usam rede e várias são regressões de defeitos que já
aconteceram — se alguma ficar vermelha, os números do relatório não valem nada.
