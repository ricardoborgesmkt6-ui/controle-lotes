# Controle de Lotes

Painel para controlar lotes de operações, com login de **administrador** e **operadores**. O painel fica no GitHub Pages e os dados ficam numa planilha do Google.

- **Operador:** tem caixa próprio, cria lotes, lança recolhimentos e novas rodadas, e vê só os dados dele. Não exclui nada.
- **Administrador:** vê todos os operadores, um por um ou somados. Cadastra, edita e desativa operadores, lança aportes e retiradas, exclui lançamentos.

Arquivos:
- `index.html` — o painel
- `Codigo.gs` — o código da planilha (vai no Apps Script)

---

## Atualizar a partir da versão sem login

1. Na planilha, abra **Extensões → Apps Script**, apague todo o código, cole o conteúdo novo do `Codigo.gs` e clique em **Salvar**.
2. Clique em **Deploy → Manage deployments**, selecione **Controle de Lotes**, clique no **lápis**, escolha **Version: New version** e clique em **Deploy**. O link do app continua o mesmo.
3. Abra o painel. Vai aparecer a tela **Criar administrador**. Preencha seu nome, usuário, senha e a **chave da planilha** (menu Controle de Lotes → Mostrar chave de acesso).
4. Pronto: os lançamentos que você já tinha passam a ser seus, como administrador.

Opcional: na planilha, rode **Controle de Lotes → Preparar planilha** de novo para esconder as colunas técnicas (IDs e senhas).

## Instalar do zero

### 1. Planilha
1. Crie uma planilha em branco (sheets.new) chamada **Controle de Lotes**.
2. **Extensões → Apps Script**: apague o código, cole o `Codigo.gs` e clique em **Salvar**.
3. Recarregue a planilha e use **Controle de Lotes → Preparar planilha**. Autorize: sua conta → **Avançado** → **Acessar (não seguro)** → **Permitir**. Isso aparece porque o código é seu, e não um app publicado.
4. Anote a **chave** que aparecer. Ela só é usada para criar o administrador.

### 2. App da Web
1. No Apps Script, clique em **Deploy → New deployment**.
2. Na engrenagem pequena ao lado de **Select type**, escolha **Web app**.
3. Preencha **Execute as: Me** e **Who has access: Anyone**. Se "Anyone" não aparecer, troque antes o "Execute as" para **Me**.
4. Clique em **Deploy** e copie o **Web app URL** (termina em `/exec`).

### 3. GitHub Pages
Em **Settings → Pages**, escolha a branch **main** e a pasta **/ (root)** e clique em **Save**. O painel fica em `https://SEU-USUARIO.github.io/controle-lotes/`.

### 4. Primeiro acesso
Abra o painel, cole o link do app, e crie o administrador com a chave.

---

## Cadastrar operadores

1. No painel, abra a aba **Operadores** e clique em **Novo operador**.
2. Informe o nome, o usuário (login), a senha inicial e o capital inicial dele.
3. O painel mostra uma mensagem pronta com o **link de acesso**, o usuário e a senha. Copie e envie para o operador.
4. O operador abre o link, entra e, de preferência, troca a senha em **Ajustes → Trocar senha**.

O **link de acesso** já leva o endereço da planilha, então o operador não precisa colar nada além do usuário e da senha. Sozinho, o link não dá acesso a nenhum dado.

## No dia a dia

- **Visão** (só admin): escolha "Todos os operadores" ou um operador específico. Os números, o gráfico, os lotes e o histórico seguem a visão escolhida.
- **Aba Operadores** (só admin): mostra capital, investido, recolhido, resultado, saldo e lotes em aberto de cada um, com o total. Toque numa linha para abrir o painel daquele operador.
- **Novo lançamento:** cria um lote. O admin escolhe para qual operador. O nome do lote não pode repetir para o mesmo operador.
- **Lançar em lote existente:** recolhimento (já recebido ou previsto) ou nova rodada no mesmo lote.
- **Caixa** (só admin): aporte ou retirada no caixa de um operador.
- **Editar operador** (só admin): muda o nome, o capital e o início, redefine a senha, transforma em admin ou desativa. Desativar bloqueia o login na hora e mantém os lançamentos.

## Segurança

- As senhas ficam guardadas na planilha só em forma protegida (hash), nunca em texto.
- Depois de 5 senhas erradas, o usuário fica bloqueado por 10 minutos.
- A sessão dura 30 dias em cada aparelho. Trocar a senha encerra as sessões antigas daquela pessoa.
- **Esqueceu a senha?** O admin redefine em Operadores → Editar → Nova senha. Se for a senha do próprio admin, use na planilha **Controle de Lotes → Redefinir senha de um usuário**.
- **Desconectar todo mundo:** na planilha, use **Controle de Lotes → Gerar nova chave**.

## Se algo der errado

| Mensagem | O que fazer |
|---|---|
| Sem conexão com a planilha | Confira a internet. No app da Web, "Who has access" precisa estar como **Anyone**. |
| A planilha respondeu de um jeito inesperado | O link precisa terminar em `/exec`. |
| Chave de acesso incorreta | Pegue a chave em Controle de Lotes → Mostrar chave de acesso. |
| Sua sessão expirou | Entre de novo com usuário e senha. |
