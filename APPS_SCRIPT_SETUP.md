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

## Caminho mais rápido: substituir o arquivo inteiro

O backend completo e atualizado está em **`apps-script/Code.gs`** neste
repositório. No editor do Apps Script, apague todo o conteúdo do `Code.gs`,
cole o arquivo daqui, salve e republique (passo 4 abaixo). É o mesmo código dos
blocos abaixo, mais alguns ajustes de robustez:

- trava (`LockService`) para duas pessoas gravando ao mesmo tempo não
  embaralharem as linhas
- data e hora gravadas como texto, para o Sheets não converter `11/08/2026`
  conforme o idioma nem transformar `20:42:07` em horário serial `1899-12-30T...`
- `updateRecord`/`deleteRecord` avisam quando o registro não é encontrado, em vez
  de responder "salvo" sem alterar nada
- `saveUsers` recusa uma lista vazia, para uma requisição com problema não apagar
  todos os colaboradores
- notas que começam com `=` não viram fórmula na planilha

## Ou editar à mão

1. Abra a planilha do ponto → menu **Extensões › Apps Script**.
2. Cole o bloco 1 dentro do `handleRequest(e)`, junto dos outros `if (action === ...)`.
3. Cole o bloco 2 no final do arquivo, fora de qualquer função.
4. Salve (Ctrl+S) e clique em **Implantar › Gerenciar implantações › editar (lápis) ›
   Versão: Nova versão › Implantar**. Sem esse passo o endereço `/exec` continua
   rodando o código antigo.
5. Recarregue o app e troque a senha em **Admin Panel › Change Password**.

## Bloco 1 — dentro do `handleRequest(e)`

O `doGet` e o `doPost` do script apenas repassam para `handleRequest(e)`, que já
tem `action` e `ss` no escopo. Cole isto logo **antes** do
`if (action === "getRecords")`, depois do bloco que cria a aba `Records`:

```javascript
    // ── SETTINGS (senha do admin + palavra de segurança) ──
    if (action === "getSettings") {
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, settings: readSettings_(ss) }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "saveSettings") {
      const incoming = JSON.parse(e.postData.contents).settings || {};
      writeSettings_(ss, incoming);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
```

## Bloco 2 — no final do arquivo

Fora de qualquer função, depois do último `}` do arquivo:

```javascript
function readSettings_(ss) {
  const out = { pwdHash: "", wordHash: "" };
  const sh = ss.getSheetByName("Settings");
  if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[0]) out[String(r[0])] = String(r[1] || "");
  });
  return out;
}

function writeSettings_(ss, settings) {
  const current = readSettings_(ss);
  const next = {
    pwdHash: settings.pwdHash ? String(settings.pwdHash) : current.pwdHash,
    wordHash: settings.wordHash ? String(settings.wordHash) : current.wordHash
  };
  let sh = ss.getSheetByName("Settings");
  if (!sh) sh = ss.insertSheet("Settings");
  sh.clear();
  sh.appendRow(["key", "value"]);
  sh.appendRow(["pwdHash", next.pwdHash]);
  sh.appendRow(["wordHash", next.wordHash]);
}
```

## Observações

- A aba `Settings` é criada sozinha na primeira gravação. Não renomeie.
- Enquanto nenhuma senha for definida, o app aceita a senha padrão `credix2026`.
  Depois da primeira troca, ela deixa de funcionar.
- Para esquecer a senha e recuperar o acesso, defina a **Security Word** em
  Admin Panel; ela também fica salva como hash na mesma aba.
- Para zerar tudo (voltar à senha padrão), apague as linhas `pwdHash` e
  `wordHash` da aba `Settings`.
