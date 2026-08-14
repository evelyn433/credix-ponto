# Backend na planilha (Google Apps Script)

O app é uma página estática. Quem guarda os dados é uma planilha do Google, e o
que conversa com ela é um Apps Script publicado como Web App.

## O que fica onde

| Aba | Conteúdo |
|---|---|
| `Records` | um registro por linha — batida de ponto ou ocorrência |
| `Users` | colaboradores, quem é admin, e o hash da senha e da palavra de segurança de cada um |
| `Requests` | pedidos de correção de ponto, com quem pediu, o motivo e quem aprovou |
| `Settings` | legado, esvaziada na migração — antes guardava a senha única do admin |

As abas são criadas automaticamente na primeira vez que forem necessárias.

## Senhas

Cada pessoa tem a sua. O que vai para a planilha é sempre o **hash SHA-256** —
a senha em texto puro nunca é gravada nem transmitida.

Os hashes **não são devolvidos ao navegador**. A conferência acontece dentro do
Apps Script, na ação `login`. Isso importa porque o endereço `/exec` é público:
se os hashes viessem na resposta, qualquer um com o link poderia baixá-los e
tentar quebrá-los offline.

Para gravar qualquer segredo o backend exige uma prova, conferida no servidor:

- a senha atual da pessoa, ou
- a palavra de segurança dela (é o "Forgot password?"), ou
- a senha de um admin (é o "Reset password" do Admin Panel)

A única exceção é o primeiro acesso: quem ainda não tem senha cria a sua ao
entrar pela primeira vez.

## Como atualizar o script

1. Abra a planilha do ponto → menu **Extensões › Apps Script**.
2. Apague todo o conteúdo do `Code.gs` e cole o arquivo **`apps-script/Code.gs`**
   deste repositório.
3. Salve (**Ctrl+S**).
4. **Implantar › Gerenciar implantações › ✏️ (lápis) › Versão: Nova versão › Implantar**.

O passo 4 é obrigatório: sem ele o endereço `/exec` continua rodando o código
antigo. E confira que o `Deployment ID` é o mesmo que está no `API_URL` do
`index.html` — é assim que se sabe que é a implantação certa.

Não altere **Execute as** (deve ser você, é o que dá permissão de escrever na
planilha) nem **Who has access** (deve ser *Anyone*, senão a página não
consegue ler o backend sem exigir login do Google de cada pessoa).

## Migração da senha antiga

Quando havia uma senha só para o admin, ela morava na aba `Settings`. Na
primeira leitura do código novo ela passa para o primeiro admin sem senha
própria, e a aba `Settings` é esvaziada — assim, com dois admins, o segundo não
herda a mesma senha. Quem não é admin começa sem senha e cria a sua no primeiro
acesso.

## Ações do backend

| Ação | Método | O que faz |
|---|---|---|
| `getRecords` | POST | devolve os registros — **exige usuário e senha válidos** |
| `addRecord` | POST | grava uma batida ou ocorrência |
| `updateRecord` | POST | corrige um registro pelo `id` |
| `deleteRecord` | POST | apaga um registro pelo `id` |
| `getUsers` | GET | colaboradores, sem hash — só se cada um já tem senha |
| `saveUsers` | POST | salva a lista, preservando as senhas de quem continua |
| `login` | POST | confere a senha e devolve a pessoa **e os registros**, poupando uma ida ao servidor |
| `setPassword` | POST | grava uma senha nova, exigindo prova |
| `setSecurityWord` | POST | grava a palavra de segurança, exigindo a senha atual |
| `resetPassword` | POST | admin zera a senha de alguém |
| `getRequests` | POST | pedidos de correção — os seus, ou todos se for admin |
| `createRequest` | POST | abre um pedido, só para si mesmo e com motivo obrigatório |
| `reviewRequest` | POST | admin aprova ou rejeita; aprovar é o que grava a batida |

`getSettings` e `saveSettings` foram removidas de propósito: eram a senha única
do admin, e uma aba com o código antigo em cache leria "nenhuma senha" e
voltaria a aceitar a senha padrão.

## Leitura dos registros exige senha

A página é pública e o endereço do backend está no código-fonte dela. Se
`getRecords` fosse aberto, qualquer um com o link leria a jornada de trabalho de
todo mundo. Então ler exige usuário e senha, conferidos aqui.

`getUsers` continua aberto de propósito: a tela de login precisa da lista de
nomes antes de alguém entrar. Isso expõe os nomes do time e quem já criou senha
— não expõe registro de ponto nem hash.

Consequência prática: não dá mais para conferir os registros colando a URL no
navegador. Para diagnóstico, `?action=getUsers` continua funcionando.

## Correções de ponto

Um colaborador não edita as próprias batidas — isso tornaria o registro sem
valor como comprovação. Ele abre um **pedido**, que fica na aba `Requests` e não
entra em nenhum cálculo de horas. Só quando o admin aprova é que este backend
grava a batida em `Records`.

Dois tipos de pedido, decididos automaticamente pela tela:

- **adicionar** — não existe batida daquele tipo naquele dia (esqueceu de bater)
- **alterar** — já existe, e o pedido carrega o `targetId` dela

Regras conferidas no servidor:

- só o próprio dono pede correção, e só sobre batidas dele
- o motivo é obrigatório
- aprovar e rejeitar exigem senha de admin — ninguém aprova o próprio pedido
- um pedido já decidido não pode ser decidido de novo
- se a batida a alterar foi apagada no meio do caminho, a aprovação falha e o
  pedido continua pendente, em vez de gravar algo errado

A linha do pedido nunca é apagada: ela é o histórico de quem pediu, por quê, e
quem aprovou.

## Custo de cada chamada

Cada ação custa leituras de planilha, e é isso que faz o app parecer lento — não
a rede. Um `login` faz 4 leituras: usuários, registros, pedidos, e a abertura da
aba `Records`. Duas coisas evitam desperdício:

- **A aba `Records` é aberta só por quem precisa dela.** `getUsers`, que roda no
  carregamento de toda página, não toca nela.
- **A migração da senha antiga é marcada numa propriedade do script.** Antes ela
  lia a aba `Settings` em toda requisição que olhasse um usuário, para descobrir
  que não havia nada a migrar.

## Cuidados que o backend toma

- **`LockService`** serializa as gravações, para duas pessoas batendo ponto no
  mesmo segundo não embaralharem as linhas
- **Tudo é gravado como texto** (com apóstrofo), senão o Sheets reinterpreta
  `11/08/2026` conforme o idioma e transforma `20:42:07` no horário serial
  `1899-12-30T...` — era o que chegava torto no app
- **`updateRecord`/`deleteRecord` avisam quando não acham o registro**, em vez de
  responder "salvo" sem alterar nada
- **`saveUsers` recusa lista vazia**, para uma requisição com problema não apagar
  todos os colaboradores
- **Notas que começam com `=` não viram fórmula** na planilha

## Emergência: ninguém consegue entrar

Abra a aba `Users` na planilha e apague o valor da coluna `pwdHash` da pessoa.
Ela cria uma senha nova no próximo acesso. Nenhum registro de ponto é afetado.

## Descobrir qual planilha está ligada ao script

No editor do Apps Script, selecione a função `qualPlanilha` e clique em **▶ Run**.
O nome e o link aparecem no painel **Registro de execução**.
