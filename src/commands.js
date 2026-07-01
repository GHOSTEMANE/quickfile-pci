/* Arquivar — menu-ao-enviar (OnMessageSend) + funcao "Enviar e arquivar".
   Corre no runtime de evento (commands.html) no Outlook Mac/web/novo Windows. */

const CLIENT_ID = "b23c4cee-bc85-4b4c-ad0a-a4422b028cda";
const AUTHORITY = "https://login.microsoftonline.com/6b1e5ee4-5f29-4a83-9c85-ac56ea7118d2";
const SCOPES = ["Mail.ReadWrite", "User.Read"];
const SISTEMA = ["caixa de entrada","inbox","rascunhos","drafts","itens enviados","enviada","sent items",
  "itens eliminados","deleted items","lixo","junk","e-mail de lixo","a enviar","outbox",
  "problemas de sincronização","histórico de conversações","conversation history","arquivo morto"];

// Primeiro teste: mostrar diagnostico no proprio aviso quando nao arquiva. Por a false depois de validar.
const DIAG = true;

let msalInstance, cachedToken = null;
const ehSistema = (n) => SISTEMA.includes((n || "").toLowerCase());

/* ---------- auth: so silencioso (runtime sem UI); usa a sessao ja iniciada no painel (mesma origem) ---------- */
async function initMsal() {
  if (!msalInstance) msalInstance = await msal.createNestablePublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: AUTHORITY }, cache: { cacheLocation: "localStorage" } });
}
async function obterToken() {
  if (cachedToken) return cachedToken;
  await initMsal();
  cachedToken = (await msalInstance.acquireTokenSilent({ scopes: SCOPES })).accessToken;
  return cachedToken;
}
async function graphFetch(url, opts) {
  const t = await obterToken();
  const o = Object.assign({}, opts, { headers: Object.assign({ Authorization: "Bearer " + t }, (opts && opts.headers) || {}) });
  return fetch(url, o);
}

/* ---------- helpers Office (promisificados) ---------- */
function destinatarios(item) { return new Promise((res) => { try { item.to.getAsync((r) => res(r.status === Office.AsyncResultStatus.Succeeded ? (r.value || []) : [])); } catch (e) { res([]); } }); }
function guardarRascunho(item) { return new Promise((res) => { try { item.saveAsync((r) => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null)); } catch (e) { res(null); } }); }
function lerSessao(item, k) { return new Promise((res) => { try { item.sessionData.getAsync(k, (r) => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null)); } catch (e) { res(null); } }); }
function escreverSessao(item, k, v) { return new Promise((res) => { try { item.sessionData.setAsync(k, v, (r) => res(r.status === Office.AsyncResultStatus.Succeeded)); } catch (e) { res(false); } }); }
function lerContexto(item) { return new Promise((res) => { try { item.getInitializationContextAsync((r) => { if (r.status === Office.AsyncResultStatus.Succeeded && r.value) { try { res(JSON.parse(r.value)); } catch (e) { res(null); } } else res(null); }); } catch (e) { res(null); } }); }

/* ---------- sugerir pasta: onde estao arquivados emails deste contacto ---------- */
async function sugerirPasta(endereco) {
  const url = "https://graph.microsoft.com/v1.0/me/messages?$top=30&$select=parentFolderId&$search=" + encodeURIComponent('"from:' + endereco + '"');
  const r = await graphFetch(url);
  if (!r.ok) throw new Error("procura " + r.status);
  const counts = {};
  ((await r.json()).value || []).forEach((m) => { if (m.parentFolderId) counts[m.parentFolderId] = (counts[m.parentFolderId] || 0) + 1; });
  const ord = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);
  for (const fid of ord.slice(0, 4)) {
    const fr = await graphFetch("https://graph.microsoft.com/v1.0/me/mailFolders/" + fid + "?$select=displayName");
    if (!fr.ok) continue;
    const nome = (await fr.json()).displayName || "";
    if (nome && !ehSistema(nome)) return { id: fid, nome: nome };
  }
  return null;
}

/* ---------- HANDLER do evento OnMessageSend ---------- */
async function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;
  try {
    // evitar loop: se ja tratamos este envio, deixar seguir
    const done = await lerSessao(item, "qf_done");
    if (done) { event.completed({ allowEvent: true }); return; }

    const to = await destinatarios(item);
    const endereco = to.length ? (to[0].emailAddress || "").toLowerCase() : "";
    if (!endereco) { event.completed({ allowEvent: true }); return; }

    let sug = null, erro = "";
    try { sug = await sugerirPasta(endereco); } catch (e) { erro = (e && e.message) || String(e); }

    if (sug) {
      event.completed({
        allowEvent: false,
        errorMessage: 'Guardar uma copia deste email em "' + sug.nome + '"?\nToca em "Enviar e arquivar" para guardar a copia e enviar, ou "Enviar na mesma" para enviar sem guardar.',
        cancelLabel: "Enviar e arquivar",
        commandId: "arquivarEnviarButton",
        contextData: JSON.stringify({ folderId: sug.id, folderName: sug.nome })
      });
    } else if (DIAG) {
      event.completed({
        allowEvent: false,
        errorMessage: erro
          ? ('Arquivar PCI (diagnostico): nao consegui sugerir pasta para ' + endereco + '. Motivo: ' + erro + '. Podes enviar na mesma.')
          : ('Arquivar PCI: ainda nao ha pasta arquivada para ' + endereco + '. Podes enviar na mesma.')
      });
    } else {
      event.completed({ allowEvent: true });
    }
  } catch (e) {
    try { event.completed({ allowEvent: true }); } catch (e2) {}
  }
}

/* ---------- FUNCAO do botao "Enviar e arquivar" (no aviso) ---------- */
async function arquivarEEnviar(event) {
  const item = Office.context.mailbox.item;
  try {
    const ctx = await lerContexto(item);
    if (ctx && ctx.folderId) {
      const id = await guardarRascunho(item);
      if (id) {
        try {
          const restId = Office.context.mailbox.convertToRestId(id, Office.MailboxEnums.RestVersion.v2_0);
          await graphFetch("https://graph.microsoft.com/v1.0/me/messages/" + restId + "/copy", {
            method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ destinationId: ctx.folderId }) });
        } catch (e) {}
      }
      await escreverSessao(item, "qf_done", "1");
    }
    // enviar programaticamente (dispara OnMessageSend, mas o loop-guard deixa passar)
    item.sendAsync((r) => { try { event.completed(); } catch (e) {} });
  } catch (e) {
    try { item.sendAsync(() => { try { event.completed(); } catch (e2) {} }); } catch (e3) { try { event.completed(); } catch (e4) {} }
  }
}

// Mapear os nomes do manifesto para as funcoes.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("arquivarEEnviar", arquivarEEnviar);
