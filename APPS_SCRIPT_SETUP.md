# Backend na planilha (Google Apps Script)

O app é uma página estática. Quem guarda os dados é uma planilha do Google, e o
que conversa com ela é um Apps Script publicado como Web App.

## O que fica onde

| Aba | Conteúdo |
|---|---|
| `Records` | um registro por linha — batida de ponto ou ocorrência |
| `Users` | colaboradores, quem é admin, e o hash da senha e da palavra de segurança de cada um |
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
| `getRecords` | GET | devolve todos os registros |
| `addRecord` | POST | grava uma batida ou ocorrência |
| `updateRecord` | POST | corrige um registro pelo `id` |
| `deleteRecord` | POST | apaga um registro pelo `id` |
| `getUsers` | GET | colaboradores, sem hash — só se cada um já tem senha |
| `saveUsers` | POST | salva a lista, preservando as senhas de quem continua |
| `login` | POST | confere a senha e devolve os dados da pessoa |
| `setPassword` | POST | grava uma senha nova, exigindo prova |
| `setSecurityWord` | POST | grava a palavra de segurança, exigindo a senha atual |
| `resetPassword` | POST | admin zera a senha de alguém |

`getSettings` e `saveSettings` foram removidas de propósito: eram a senha única
do admin, e uma aba com o código antigo em cache leria "nenhuma senha" e
voltaria a aceitar a senha padrão.

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
