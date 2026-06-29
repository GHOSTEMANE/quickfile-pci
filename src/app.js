/* Arquivar — Fases 4/5/6: arquivar + sugestao + criar pasta. Rapido (cache pastas, sessao). */

const CLIENT_ID = "b23c4cee-bc85-4b4c-ad0a-a4422b028cda";
const AUTHORITY = "https://login.microsoftonline.com/6b1e5ee4-5f29-4a83-9c85-ac56ea7118d2";
const SCOPES = ["Mail.ReadWrite", "User.Read"];
const LS_PASTAS = "qf_pastas_v1";
const LS_USO = "qf_uso_v1";
const SISTEMA = ["caixa de entrada","inbox","rascunhos","drafts","itens enviados","enviada","sent items",
  "itens eliminados","deleted items","lixo","junk","e-mail de lixo","a enviar","outbox",
  "problemas de sincronização","histórico de conversações","conversation history","arquivo morto"];

let msalInstance, cachedToken = null;
let pastas = [], pastasById = {};
let emailAtual = { remetente: "", itemId: null };
let uso = JSON.parse(localStorage.getItem(LS_USO) || "{}");

const $ = (id) => document.getElementById(id);
function status(msg, cls) { const s = $("status"); s.textContent = msg || ""; s.className = cls || ""; }
const ehSistema = (n) => SISTEMA.includes((n || "").toLowerCase());

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) { status("Este add-in destina-se ao Outlook."); return; }
  $("busca").addEventListener("input", () => renderLista());
  $("btnLogin").addEventListener("click", () => bootstrap(true));
  $("novaPasta").addEventListener("keydown", (ev) => { if (ev.key === "Enter") criarPasta($("novaPasta").value); });
  try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, atualizarSeMudou); } catch (e) {}
  setInterval(atualizarSeMudou, 700); // segue a troca de email
  // apanha pastas criadas fora (no Outlook) ao voltar ao painel
  document.addEventListener("visibilitychange", () => { if (!document.hidden) refreshPastas().catch(() => {}); });

  renderEmail();
  const cache = JSON.parse(localStorage.getItem(LS_PASTAS) || "null");
  if (cache && cache.length) { setPastas(cache); mostrarUI(); calcularSugestao(); }
  bootstrap(false);
});

/* ---------- email selecionado ---------- */
function renderEmail() {
  const item = Office.context.mailbox.item;
  const el = $("email");
  if (item && item.subject) {
    const rem = (item.from && item.from.emailAddress) || "";
    const nome = (item.from && item.from.displayName) || rem;
    emailAtual = { remetente: rem.toLowerCase(), itemId: item.itemId };
    el.className = "card"; el.innerHTML = "";
    const a = document.createElement("div"); a.className = "assunto"; a.textContent = item.subject;
    const d = document.createElement("div"); d.className = "de"; d.textContent = nome ? ("de " + nome) : "";
    el.appendChild(a); el.appendChild(d);
  } else {
    emailAtual = { remetente: "", itemId: null };
    el.className = "card vazio"; el.textContent = "Seleciona um email na lista.";
  }
}
function atualizarSeMudou() {
  const item = Office.context.mailbox.item;
  const id = item ? item.itemId : null;
  if (id !== emailAtual.itemId) { renderEmail(); calcularSugestao(); }
}

/* ---------- auth ---------- */
async function initMsal() {
  if (!msalInstance) {
    msalInstance = await msal.createNestablePublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: AUTHORITY }, cache: { cacheLocation: "localStorage" },
    });
  }
}
async function ensureToken(allowPopup) {
  if (cachedToken) return cachedToken;
  await initMsal();
  try { cachedToken = (await msalInstance.acquireTokenSilent({ scopes: SCOPES })).accessToken; }
  catch (e) { if (!allowPopup) throw e; cachedToken = (await msalInstance.acquireTokenPopup({ scopes: SCOPES })).accessToken; }
  return cachedToken;
}
async function graphFetch(url, opts) {
  const build = (tok) => Object.assign({}, opts, { headers: Object.assign({ Authorization: "Bearer " + tok }, (opts && opts.headers) || {}) });
  let t = await ensureToken(true);
  let r = await fetch(url, build(t));
  if (r.status === 401) { cachedToken = null; t = await ensureToken(true); r = await fetch(url, build(t)); }
  return r;
}

/* ---------- bootstrap / pastas ---------- */
async function bootstrap(interactive) {
  try {
    await ensureToken(interactive);
    await refreshPastas();
    mostrarUI();
    calcularSugestao();
    $("btnLogin").classList.add("hidden");
  } catch (e) {
    if (!pastas.length) status("Liga a tua conta para começar.");
    $("btnLogin").classList.remove("hidden");
  }
}
async function refreshPastas() {
  const url = "https://graph.microsoft.com/v1.0/me/mailFolders?$top=200&$select=id,displayName,totalItemCount";
  const r = await graphFetch(url);
  if (!r.ok) return;
  const data = await r.json();
  const lista = (data.value || []).sort((a, b) => (a.displayName || "").localeCompare(b.displayName || "", "pt"));
  setPastas(lista);
  localStorage.setItem(LS_PASTAS, JSON.stringify(lista));
}
function setPastas(lista) {
  pastas = lista; pastasById = {};
  lista.forEach((p) => { pastasById[p.id] = p; });
  renderLista();
}
function mostrarUI() { $("busca").classList.remove("hidden"); $("novaWrap").classList.remove("hidden"); }

/* ---------- render lista ---------- */
function renderLista() {
  const q = $("busca").value.trim().toLowerCase();
  const ul = $("lista"); ul.innerHTML = "";
  pastas.filter((p) => !q || (p.displayName || "").toLowerCase().includes(q)).forEach((p) => ul.appendChild(itemPasta(p)));
}
function itemPasta(p, sugerida) {
  const li = document.createElement("li");
  if (sugerida) li.className = "sug";
  const nome = document.createElement("span"); nome.className = "nome"; nome.textContent = p.displayName || "(sem nome)";
  const cnt = document.createElement("span"); cnt.className = "cnt";
  cnt.textContent = typeof p.totalItemCount === "number" ? String(p.totalItemCount) : "";
  li.appendChild(nome); li.appendChild(cnt);
  li.addEventListener("click", () => arquivar(p.id, p.displayName));
  return li;
}

/* ---------- sugestao ---------- */
function calcularSugestao() {
  $("sugerido").innerHTML = ""; $("sugWrap").classList.add("hidden");
  const rem = emailAtual.remetente;
  if (!rem || !pastas.length) return;
  const local = melhorDoUso(rem);
  if (local && pastasById[local]) { mostrarSugerida(local); return; }
  descobrirSugestao(rem);
}
function melhorDoUso(rem) {
  const m = uso[rem]; if (!m) return null;
  let best = null, bestScore = 0;
  Object.keys(m).forEach((fid) => { if (m[fid] > bestScore && pastasById[fid] && !ehSistema(pastasById[fid].displayName)) { best = fid; bestScore = m[fid]; } });
  return best;
}
async function descobrirSugestao(rem) {
  try {
    const url = "https://graph.microsoft.com/v1.0/me/messages?$top=25&$select=parentFolderId&$search=" + encodeURIComponent('"from:' + rem + '"');
    const r = await graphFetch(url);
    if (!r.ok) return;
    const data = await r.json();
    const counts = {};
    (data.value || []).forEach((m) => { const f = m.parentFolderId; if (pastasById[f] && !ehSistema(pastasById[f].displayName)) counts[f] = (counts[f] || 0) + 1; });
    let best = null, bestN = 0;
    Object.keys(counts).forEach((f) => { if (counts[f] > bestN) { best = f; bestN = counts[f]; } });
    if (best && emailAtual.remetente === rem) mostrarSugerida(best);
  } catch (e) {}
}
function mostrarSugerida(folderId) {
  const p = pastasById[folderId]; if (!p) return;
  $("sugerido").innerHTML = ""; $("sugerido").appendChild(itemPasta(p, true));
  $("sugWrap").classList.remove("hidden");
}

/* ---------- arquivar (+ aprende, + atualiza contagem) ---------- */
async function arquivar(folderId, folderName) {
  const item = Office.context.mailbox.item;
  if (!item || !item.itemId) { status("Seleciona primeiro um email.", "err"); return; }
  const rem = emailAtual.remetente;
  status('A arquivar em "' + folderName + '"…');
  try {
    const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    const r = await graphFetch("https://graph.microsoft.com/v1.0/me/messages/" + restId + "/move", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ destinationId: folderId }),
    });
    if (r.ok) {
      aprender(rem, folderId);
      if (pastasById[folderId]) { pastasById[folderId].totalItemCount = (pastasById[folderId].totalItemCount || 0) + 1; renderLista(); }
      status('✅ Arquivado em "' + folderName + '".', "ok");
      $("email").className = "card vazio"; $("email").textContent = "Arquivado. Seleciona outro email.";
      $("sugWrap").classList.add("hidden");
      refreshPastas().catch(() => {}); // sincroniza contagens com o servidor
    } else {
      status("Não deu para arquivar (" + r.status + "): " + (await r.text()).slice(0, 250), "err");
    }
  } catch (e) { status("Erro ao arquivar: " + ((e && e.message) || e), "err"); }
}
function aprender(rem, folderId) {
  if (!rem) return;
  uso[rem] = uso[rem] || {}; uso[rem][folderId] = (uso[rem][folderId] || 0) + 1;
  localStorage.setItem(LS_USO, JSON.stringify(uso));
}

/* ---------- criar pasta (Fase 6) ---------- */
async function criarPasta(nome) {
  nome = (nome || "").trim();
  if (!nome) return;
  status('A criar a pasta "' + nome + '"…');
  try {
    const r = await graphFetch("https://graph.microsoft.com/v1.0/me/mailFolders", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ displayName: nome }),
    });
    if (!r.ok) { status("Não deu para criar (" + r.status + "): " + (await r.text()).slice(0, 250), "err"); return; }
    const nova = await r.json();
    pastas.push({ id: nova.id, displayName: nova.displayName, totalItemCount: nova.totalItemCount || 0 });
    pastas.sort((a, b) => (a.displayName || "").localeCompare(b.displayName || "", "pt"));
    setPastas(pastas);
    localStorage.setItem(LS_PASTAS, JSON.stringify(pastas));
    $("novaPasta").value = "";
    if (emailAtual.itemId) { await arquivar(nova.id, nova.displayName); }
    else { status('✅ Pasta "' + nova.displayName + '" criada.', "ok"); }
  } catch (e) { status("Erro ao criar pasta: " + ((e && e.message) || e), "err"); }
}
