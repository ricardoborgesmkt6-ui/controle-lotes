# Controle de Lotes — como colocar no ar

São 3 partes, feitas uma vez só: a **planilha** (onde os dados ficam), o **app da Web** (a ponte entre a planilha e o painel) e o **GitHub Pages** (onde o painel fica hospedado). Leva uns 15 minutos.

Arquivos:
- `index.html` — o painel (vai para o GitHub)
- `Codigo.gs` — o código da planilha (vai no Apps Script)

---

## Parte 1 — Planilha

1. No Google Drive, crie uma **planilha em branco** e dê o nome **Controle de Lotes**.
2. Na planilha, abra **Extensões → Apps Script**.
3. Apague todo o código que aparece, cole o conteúdo do arquivo `Codigo.gs` e clique em **Salvar** (ícone de disquete).
4. Volte para a aba da planilha e **recarregue a página** (F5). Vai aparecer o menu **Controle de Lotes** no topo (pode levar alguns segundos).
5. Clique em **Controle de Lotes → Preparar planilha**.
   - O Google vai pedir autorização. Escolha sua conta → **Avançado** → **Acessar Controle de Lotes (não seguro)** → **Permitir**. Esse aviso aparece porque o código é seu, e não de um app publicado; é normal.
   - Depois de autorizar, clique de novo em **Preparar planilha** se a mensagem não aparecer.
6. Vai aparecer a sua **chave de acesso**. Copie e guarde. Se perder, use **Controle de Lotes → Mostrar chave de acesso**.

A planilha fica com 4 abas:
- **Resumo**: saldo em caixa, investido, recolhido, resultado e retorno, tudo por fórmula.
- **Lançamentos**: cada lançamento é uma linha (data, lote, tipo, quantidade, valor, situação).
- **Lotes**: cada lote, com investido, recolhido, a recolher, resultado e situação calculados.
- **Config**: capital inicial e data de início.

Não edite as colunas de ID (ficam ocultas) nem as colunas calculadas. Para lançar, use o painel.

## Parte 2 — App da Web (a ponte)

1. No Apps Script, clique em **Implantar → Nova implantação**.
2. Na engrenagem ao lado de "Selecione o tipo", escolha **App da Web**.
3. Preencha:
   - Descrição: `Controle de Lotes`
   - **Executar como: Eu**
   - **Quem pode acessar: Qualquer pessoa**
4. Clique em **Implantar** e copie o **URL do app da Web** (termina em `/exec`).

"Qualquer pessoa" é necessário para o painel conseguir falar com a planilha. Sem a sua chave, ninguém lê nem grava nada.

## Parte 3 — GitHub Pages

1. No GitHub, crie um repositório novo, por exemplo `controle-lotes`, marcado como **Public**.
2. Clique em **Add file → Upload files**, envie o `index.html` e confirme em **Commit changes**.
3. Vá em **Settings → Pages**. Em "Branch", escolha **main** e a pasta **/ (root)** e clique em **Save**.
4. Em 1 ou 2 minutos o painel fica no ar em: `https://SEU-USUARIO.github.io/controle-lotes/`

O repositório público mostra só o código do painel. Os seus dados ficam na planilha, e a chave fica guardada só no seu aparelho.

## Parte 4 — Conectar

1. Abra o link do painel.
2. Cole o **URL do app da Web** e a **chave de acesso** e clique em **Conectar**.
3. Informe o **capital inicial** e a **data de início**. Pronto.

Faça a conexão uma vez em cada aparelho (celular e computador). No celular, use "Adicionar à tela inicial" para abrir como um app.

---

## Uso do dia a dia

- **Novo lançamento**: cria um lote com o nome que você escolher (o nome não pode repetir).
- **Lançar em lote existente**: escolha o lote e lance um **Recolhimento** (já recebido ou previsto) ou uma **Nova rodada** no mesmo lote.
- **Caixa**: aportes e retiradas que não são de nenhum lote.
- **Recebido**: num recolhimento previsto, marca que o dinheiro entrou.
- **Atualizar**: puxa os dados mais recentes da planilha. O painel também atualiza sozinho a cada minuto e quando você volta para ele.
- **Ajustes**: muda o capital inicial, abre a planilha, exporta CSV ou troca a conexão.

## Se algo der errado

| Mensagem | O que fazer |
|---|---|
| Chave de acesso incorreta | Confira a chave em Controle de Lotes → Mostrar chave de acesso. |
| A planilha respondeu de um jeito inesperado | O link precisa terminar em `/exec`, e o acesso tem que estar como **Qualquer pessoa**. |
| A planilha ainda não foi preparada | Rode Controle de Lotes → Preparar planilha. |
| Sem conexão com a planilha | Verifique a internet; tente **Atualizar**. |

**Se precisar trocar o código da planilha** (uma versão nova do `Codigo.gs`): cole o código novo, salve e vá em **Implantar → Gerenciar implantações → lápis (editar) → Versão: Nova versão → Implantar**. Assim o link continua o mesmo e não precisa reconectar os aparelhos.

**Se a chave vazar**: use Controle de Lotes → Gerar nova chave e reconecte seus aparelhos com ela.
