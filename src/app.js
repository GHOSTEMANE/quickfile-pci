/* Arquivar — Fase 4: mostrar o email selecionado + arquivar numa pasta (Graph move).
   Auth por NAA (MSAL). Cores fixas (tema escuro) no CSS. */

const CLIENT_ID = "b23c4cee-bc85-4b4c-ad0a-a4422b028cda";
const AUTHORITY = "https://login.microsoftonline.com/6b1e5ee4-5f29-4a83-9c85-ac56ea7118d2";
const SCOPES = ["Mail.ReadWrite", "User.Read"];

let msalInstance;
let pastas = null; // cache: [{id, displayName, totalItemCount}]

const $ = (id) => document.getElementById(id);
function status(msg, cls) { const s = $("status"); s.textContent = msg; s.className = cls || ""; }

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) { status("Este add-in destina-se ao Outlook."); return; }
  renderEmail();
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, renderEmail);
  } catch (e) { /* sem suporte a ItemChanged: ignora */ }
  $("btnCarregar").addEventListener("click", () => carregarPastas(false));
  $("busca").addEventListener("input", filtrar);
  // Tenta carregar as pastas em silencio (se ja houver sessao).
  carregarPastas(true);
});

/* ---------- Email selecionado ---------- */
function renderEmail() {
  const el = $("email");
  const item = Office.context.mailbox.item;
  if (item && item.subject) {
    const de = (item.from && (item.from.displayName || item.from.emailAddress)) || "";
    el.className = "card";
    el.innerHTML = "";
    const a = document.createElement("div"); a.className = "assunto"; a.textContent = item.subject;
    const d = document.createElement("div"); d.className = "de"; d.textContent = de ? ("de " + de) : "";
    el.appendChild(a); el.appendChild(d);
  } else {
    el.className = "card vazio";
    el.textContent = "Seleciona um email na lista.";
  }
}

/* ---------- Autenticacao (NAA) ---------- */
async function initMsal() {
  if (!msalInstance) {
    msalInstance = await msal.createNestablePublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: AUTHORITY },
      cache: { cacheLocation: "localStorage" },
    });
  }
}
async function getToken(allowPopup) {
  await initMsal();
  const req = { scopes: SCOPES };
  try {
    return (await msalInstance.acquireTokenSilent(req)).accessToken;
  } catch (e) {
    if (!allowPopup) throw e;
    return (await msalInstance.acquireTokenPopup(req)).accessToken;
  }
}

/* ---------- Pastas ---------- */
async function carregarPastas(silent) {
  status("A ligar à tua conta…");
  try {
    const token = await getToken(!silent);
    status("A obter as pastas…");
    const url = "https://graph.microsoft.com/v1.0/me/mailFolders" +
                "?$top=200&$select=id,displayName,totalItemCount";
    const resp = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    if (!resp.ok) { status("Erro do Graph (" + resp.status + "): " + (await resp.text()).slice(0, 300), "err"); return; }
    const data = await resp.json();
    pastas = (data.value || []).sort((a, b) => (a.displayName || "").localeCompare(b.displayName || "", "pt"));
    $("btnCarregar").classList.add("hidden");
    $("busca").classList.remove("hidden");
    $("hint").classList.remove("hidden");
    status("Pronto — escolhe a pasta.", "ok");
    renderPastas(pastas);
  } catch (e) {
    if (silent) { status("Carrega no botão para ligar a tua conta."); }
    else { status("Não consegui ligar: " + ((e && e.message) || e), "err"); }
  }
}

function renderPastas(lista) {
  const ul = $("lista");
  ul.innerHTML = "";
  lista.forEach((p) => {
    const li = document.createElement("li");
    const nome = document.createElement("span"); nome.className = "nome"; nome.textContent = p.displayName || "(sem nome)";
    const cnt = document.createElement("span"); cnt.className = "cnt";
    cnt.textContent = typeof p.totalItemCount === "number" ? String(p.totalItemCount) : "";
    li.appendChild(nome); li.appendChild(cnt);
    li.addEventListener("click", () => arquivar(p.id, p.displayName));
    ul.appendChild(li);
  });
}

function filtrar() {
  if (!pastas) return;
  const q = $("busca").value.trim().toLowerCase();
  renderPastas(q ? pastas.filter((p) => (p.displayName || "").toLowerCase().includes(q)) : pastas);
}

/* ---------- Arquivar ---------- */
async function arquivar(folderId, folderName) {
  const item = Office.context.mailbox.item;
  if (!item || !item.itemId) { status("Seleciona primeiro um email na lista.", "err"); return; }
  status('A arquivar em "' + folderName + '"…');
  try {
    const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    const token = await getToken(true);
    const resp = await fetch("https://graph.microsoft.com/v1.0/me/messages/" + restId + "/move", {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ destinationId: folderId }),
    });
    if (resp.ok) {
      status('✅ Arquivado em "' + folderName + '".', "ok");
      $("email").className = "card vazio";
      $("email").textContent = "Email arquivado. Seleciona outro.";
    } else {
      status("Não deu para arquivar (" + resp.status + "): " + (await resp.text()).slice(0, 300), "err");
    }
  } catch (e) {
    status("Erro ao arquivar: " + ((e && e.message) || e), "err");
  }
}
