/* Arquivar — Fase 3: autenticacao (NAA) + listar pastas via Microsoft Graph */

const CLIENT_ID = "b23c4cee-bc85-4b4c-ad0a-a4422b028cda";
const AUTHORITY = "https://login.microsoftonline.com/6b1e5ee4-5f29-4a83-9c85-ac56ea7118d2";
const SCOPES = ["Mail.ReadWrite", "User.Read"];

let msalInstance;

Office.onReady((info) => {
  applyTheme();
  const statusEl = document.getElementById("status");
  if (info.host !== Office.HostType.Outlook) {
    statusEl.textContent = "Este add-in destina-se ao Outlook.";
    return;
  }
  statusEl.textContent = "Add-in a funcionar ✅ — carrega no botão para ver as tuas pastas.";
  document.getElementById("btnPastas").addEventListener("click", verPastas);
});

/** Decide claro/escuro de forma robusta e FORCA um par de cores com contraste garantido.
 *  (Nao confia no bodyForegroundColor do Outlook, que vem trocado no modo escuro do Mac.) */
function applyTheme() {
  let dark = true; // o default do CSS ja e escuro
  try {
    const bg = Office.context && Office.context.officeTheme && Office.context.officeTheme.bodyBackgroundColor;
    if (bg) dark = isDarkColor(bg);
  } catch (e) { /* usa fallback abaixo */ }
  try {
    if (window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches) {
      // se o sistema diz explicitamente claro e o Office nao contrariou, respeita
      if (!(Office.context && Office.context.officeTheme && Office.context.officeTheme.bodyBackgroundColor)) {
        dark = false;
      }
    }
  } catch (e) { /* ignora */ }

  const s = document.documentElement.style;
  if (dark) {
    s.setProperty("--bg", "#1f1f1f");
    s.setProperty("--fg", "#ffffff");
    s.setProperty("--muted", "#c8c8c8");
    s.setProperty("--panel", "#383838");
    s.setProperty("--azul", "#2b88d8");
  } else {
    s.setProperty("--bg", "#ffffff");
    s.setProperty("--fg", "#202020");
    s.setProperty("--muted", "#666666");
    s.setProperty("--panel", "#f3f6fb");
    s.setProperty("--azul", "#0f6cbd");
  }
}

function isDarkColor(c) {
  c = String(c).trim();
  let r, g, b;
  if (c[0] === "#") {
    let h = c.slice(1);
    if (h.length === 3) h = h.split("").map((x) => x + x).join("");
    r = parseInt(h.slice(0, 2), 16); g = parseInt(h.slice(2, 4), 16); b = parseInt(h.slice(4, 6), 16);
  } else {
    const m = c.match(/\d+/g);
    if (!m || m.length < 3) return true;
    r = +m[0]; g = +m[1]; b = +m[2];
  }
  if ([r, g, b].some((v) => isNaN(v))) return true;
  return (0.299 * r + 0.587 * g + 0.114 * b) < 140;
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
