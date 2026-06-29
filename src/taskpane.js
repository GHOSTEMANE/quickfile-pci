/* Arquivar — Fase 3: autenticacao (NAA) + listar pastas via Microsoft Graph */

const CLIENT_ID = "b23c4cee-bc85-4b4c-ad0a-a4422b028cda";
const AUTHORITY = "https://login.microsoftonline.com/6b1e5ee4-5f29-4a83-9c85-ac56ea7118d2";
const SCOPES = ["Mail.ReadWrite", "User.Read"];

let msalInstance;

Office.onReady((info) => {
  applyOfficeTheme();
  const statusEl = document.getElementById("status");
  if (info.host !== Office.HostType.Outlook) {
    statusEl.textContent = "Este add-in destina-se ao Outlook.";
    return;
  }
  statusEl.textContent = "Add-in a funcionar ✅ — carrega no botão para ver as tuas pastas.";
  document.getElementById("btnPastas").addEventListener("click", verPastas);
});

/** Alinha as cores do painel com o tema atual do Outlook (claro/escuro). */
function applyOfficeTheme() {
  try {
    const t = Office.context && Office.context.officeTheme;
    if (!t) return;
    const s = document.documentElement.style;
    if (t.bodyBackgroundColor) s.setProperty("--bg", t.bodyBackgroundColor);
    if (t.bodyForegroundColor) s.setProperty("--fg", t.bodyForegroundColor);
    if (t.controlBackgroundColor) s.setProperty("--panel", t.controlBackgroundColor);
  } catch (e) { /* mantem o tema do CSS */ }
}

async function initMsal() {
  if (!msalInstance) {
    msalInstance = await msal.createNestablePublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: AUTHORITY },
      cache: { cacheLocation: "localStorage" },
    });
  }
}

/** Obtem um token do Graph: silencioso e, se preciso, com janela de login. */
async function getToken() {
  await initMsal();
  const req = { scopes: SCOPES };
  try {
    const r = await msalInstance.acquireTokenSilent(req);
    return r.accessToken;
  } catch (silentError) {
    const r = await msalInstance.acquireTokenPopup(req);
    return r.accessToken;
  }
}

async function verPastas() {
  const statusEl = document.getElementById("status");
  const contentEl = document.getElementById("content");
  contentEl.className = "";
  contentEl.innerHTML = "";
  statusEl.textContent = "A autenticar…";

  try {
    const token = await getToken();
    statusEl.textContent = "A obter as pastas…";

    const url = "https://graph.microsoft.com/v1.0/me/mailFolders" +
                "?$top=200&$select=displayName,totalItemCount";
    const resp = await fetch(url, { headers: { Authorization: "Bearer " + token } });

    if (!resp.ok) {
      const txt = await resp.text();
      statusEl.textContent = "Erro do Microsoft Graph (" + resp.status + ").";
      contentEl.className = "err";
      contentEl.textContent = txt.slice(0, 500);
      return;
    }

    const data = await resp.json();
    const pastas = data.value || [];
    statusEl.textContent = "✅ Encontrei " + pastas.length + " pastas na tua conta:";

    const ul = document.createElement("ul");
    pastas
      .sort((a, b) => (a.displayName || "").localeCompare(b.displayName || "", "pt"))
      .forEach((p) => {
        const li = document.createElement("li");
        const nome = document.createElement("span");
        nome.textContent = p.displayName || "(sem nome)";
        const cnt = document.createElement("span");
        cnt.className = "cnt";
        cnt.textContent = typeof p.totalItemCount === "number" ? String(p.totalItemCount) : "";
        li.appendChild(nome);
        li.appendChild(cnt);
        ul.appendChild(li);
      });
    contentEl.appendChild(ul);
  } catch (e) {
    statusEl.textContent = "Não consegui autenticar.";
    contentEl.className = "err";
    contentEl.textContent = String((e && e.message) ? e.message : e);
  }
}
