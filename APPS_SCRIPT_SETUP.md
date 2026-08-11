# Senha compartilhada entre dispositivos — ajuste no Apps Script

Hoje o app já tenta ler e gravar a senha de admin e a palavra de segurança na
planilha. Enquanto o Apps Script não responder às ações `getSettings` e
`saveSettings`, ele cai automaticamente no navegador e avisa na tela:

> ⚠️ Saved in this browser only — set up the spreadsheet backend to share it across devices.

Depois de aplicar o passo a passo abaixo, a mensagem passa a ser
"✅ Password updated on every device!" e a senha vale em qualquer celular ou
computador.

As senhas **não** são gravadas em texto puro: o navegador envia apenas um hash
SHA-256, então nem quem abre a planilha consegue ler a senha.

## Passo a passo

1. Abra a planilha do ponto → menu **Extensões › Apps Script**.
2. Cole os blocos abaixo no final do arquivo do script.
3. Adicione as duas linhas indicadas dentro do `doGet` e do `doPost` que já existem.
4. Clique em **Implantar › Gerenciar implantações › editar (lápis) › Versão: Nova versão › Implantar**.
   Sem esse passo o endereço `/exec` continua rodando o código antigo.
5. Recarregue o app e troque a senha em **Admin Panel › Change Password**.

## 1. Cole no final do script

```javascript
function getSettingsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Settings');
  if (!sh) {
    sh = ss.insertSheet('Settings');
    sh.appendRow(['key', 'value']);
  }
  return sh;
}

function readSettings_() {
  var rows = getSettingsSheet_().getDataRange().getValues().slice(1);
  var out = { pwdHash: '', wordHash: '' };
  rows.forEach(function (r) {
    if (r[0]) out[String(r[0])] = String(r[1] || '');
  });
  return out;
}

function writeSettings_(settings) {
  var sh = getSettingsSheet_();
  var current = readSettings_();
  var next = {
    pwdHash: settings && settings.pwdHash ? String(settings.pwdHash) : current.pwdHash,
    wordHash: settings && settings.wordHash ? String(settings.wordHash) : current.wordHash
  };
  sh.clear();
  sh.appendRow(['key', 'value']);
  sh.appendRow(['pwdHash', next.pwdHash]);
  sh.appendRow(['wordHash', next.wordHash]);
}

function settingsJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
```

## 2. Dentro do `doGet` que já existe

Logo no começo, antes dos outros `if`/`switch` de `action`:

```javascript
  if (action === 'getSettings') {
    return settingsJson_({ success: true, settings: readSettings_() });
  }
```

Se o seu `doGet` lê o parâmetro com outro nome, use o mesmo nome que já está lá
(normalmente `var action = e.parameter.action;`).

## 3. Dentro do `doPost` que já existe

Logo depois de fazer o parse do corpo da requisição:

```javascript
  if (body.action === 'saveSettings') {
    writeSettings_(body.settings);
    return settingsJson_({ success: true });
  }
```

Use o mesmo nome de variável que o seu `doPost` já usa para o corpo — em geral
`var body = JSON.parse(e.postData.contents);`.

## Observações

- A aba `Settings` é criada sozinha na primeira gravação. Não renomeie.
- Enquanto nenhuma senha for definida, o app aceita a senha padrão `credix2026`.
  Depois da primeira troca, ela deixa de funcionar.
- Para esquecer a senha e recuperar o acesso, defina a **Security Word** em
  Admin Panel; ela também fica salva como hash na mesma aba.
- Para zerar tudo (voltar à senha padrão), apague as linhas `pwdHash` e
  `wordHash` da aba `Settings`.
