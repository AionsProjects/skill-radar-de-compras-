# Radar de Compras no Claude Desktop

Pesquisa quanto os produtos da sua planilha custaram de verdade no comércio,
pelos portais oficiais de nota fiscal (NFC-e) do **Amazonas** e da **Paraíba**.

Os preços vêm de nota fiscal já emitida: é o que alguém realmente pagou, não
uma oferta anunciada.

---

## Antes de começar

Você precisa de **Windows** e do **Claude Desktop** já instalado.

Não precisa saber programar, nem instalar Python antes: o instalador resolve.

---

## Instalação

1. Descompacte a pasta `busca-preco` em **qualquer lugar** do seu computador.

2. Clique com o botão direito na pasta, escolha **Copiar como caminho**.

3. Abra o **PowerShell** (tecla Windows, digite `powershell`, Enter).

4. Cole o comando abaixo, trocando `CAMINHO` pelo que você copiou:

   ```powershell
   cd CAMINHO
   powershell -ExecutionPolicy Bypass -File instalar.ps1
   ```

5. Espere terminar. Ele vai:
   - procurar um Python que sirva, e oferecer instalar um se não houver;
   - instalar o que a ferramenta usa;
   - registrar a ferramenta no Claude Desktop;
   - conferir se está tudo funcionando.

6. **Feche e abra o Claude Desktop.**

---

## Como usar

Converse normalmente. Não há comando a decorar:

> cota a planilha C:\Users\eu\Documents\lista.xlsx em Manaus

> quanto custa água sanitária em Manaus?

> confere se o radar de compras está funcionando

Ao cotar uma planilha, ele grava um **PDF** e uma **planilha anotada** na mesma
pasta do arquivo original, e diz onde ficaram.

---

## Como deve ser a sua planilha

Só uma coluna é obrigatória: **Produto**.

| Coluna | Precisa? | Para que serve |
|---|---|---|
| `Produto` | sim | o nome que vai ser procurado no portal |
| `Preço Atual` | não | sem ela, o relatório vira uma cotação de mercado, sem cálculo de economia |
| `Fornecedor Atual` | não | com ela, o relatório também confere o que o seu próprio fornecedor anda cobrando |

Quanto mais completo o nome, melhor a busca: `DETERGENTE YPE NEUTRO 500ML`
acha mais do que `detergente`. Marca e tamanho ajudam.

Se a sua marca não tiver vendido no período, a busca é ampliada para a
categoria automaticamente — e o relatório avisa, na linha do item, que foi isso
que aconteceu.

Há planilhas de exemplo em `referencias/`.

---

## O que a ferramenta não faz

- **Não recomenda trocar de fornecedor.** Mostra preço, distância e origem lado
  a lado; a decisão é sua.
- **Não inventa preço.** Produto sem venda no período aparece como "sem
  conclusão", com o motivo — nunca com um valor estimado.
- **Não inclui frete nem diferença de ICMS**, e a distância é em linha reta.
  No Amazonas, confirme o acesso: muitos municípios só têm ligação fluvial.
- **Não garante o preço.** A loja não é obrigada a manter o valor, e o produto
  pode estar sem estoque.

---

## Estados

**Amazonas** funciona sozinho, sem nenhuma configuração.

**Paraíba** exige a extensão do Claude no Chrome (https://claude.ai/chrome),
porque o portal de lá pede uma verificação anti-robô que só uma pessoa resolve.

Outros estados ainda não têm suporte.

---

## Se der problema

Peça no próprio Claude Desktop:

> confere se o radar de compras está funcionando

Ele roda a bateria de testes e diz se o problema é na instalação ou no portal
(que sai do ar de vez em quando, por ser serviço público).

Se a resposta for que o Claude Desktop não conhece a ferramenta, confira se
você **fechou e abriu** o aplicativo depois de instalar.
