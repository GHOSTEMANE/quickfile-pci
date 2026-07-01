/* Arquivar — arvore de pastas (subpastas, expandir, criar subpasta), arquivar + sugestao. */

const CLIENT_ID = "b23c4cee-bc85-4b4c-ad0a-a4422b028cda";
const AUTHORITY = "https://login.microsoftonline.com/6b1e5ee4-5f29-4a83-9c85-ac56ea7118d2";
const SCOPES = ["Mail.ReadWrite", "User.Read"];
const LS_PASTAS = "qf_arvore_v1";
const LS_USO = "qf_uso_v1";
const SISTEMA = ["caixa de entrada","inbox","rascunhos","drafts","itens enviados","enviada","sent items",
  "itens eliminados","deleted items","lixo","junk","e-mail de lixo","a enviar","outbox",
  "problemas de sincronização","histórico de conversações","conversation history","arquivo morto"];

let msalInstance, cachedToken = null;
let arvore = [];            // raiz (cada no: {id, displayName, totalItemCount, childFolderCount, level, parentId, children:[]})
let pastasById = {};        // todas as pastas por id
let expandido = new Set();
let emailAtual = { itemId: null, subject: "", remetente: "", remetenteNome: "" };
let selecao = [];           // [{itemId, subject}] quando ha varios emails selecionados (>1)
let selecaoSig = "";        // assinatura da selecao multipla (evita recalcular a cada polling)
let modo = "read";          // "compose" quando o painel abre ao enviar (menu-ao-enviar)
let uso = JSON.parse(localStorage.getItem(LS_USO) || "{}");

const $ = (id) => document.getElementById(id);
function status(msg, cls) { const s = $("status"); s.textContent = msg || ""; s.className = cls || ""; }
const ehSistema = (n) => SISTEMA.includes((n || "").toLowerCase());

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) { status("Este add-in destina-se ao Outlook."); return; }
  $("busca").addEventListener("input", renderArvore);
  $("busca").addEventListener("focus", () => carregarArvore().catch(() => {}));
  $("btnLogin").addEventListener("click", () => bootstrap(true));

  // Modo compose = painel aberto ao enviar (o item tem saveAsync). Caso contrario, leitura.
  try { if (Office.context.mailbox.item && typeof Office.context.mailbox.item.saveAsync === "function") modo = "compose"; } catch (e) {}

  const cache = JSON.parse(localStorage.getItem(LS_PASTAS) || "null");
  if (cache && cache.length) { indexar(cache); arvore = cache; mostrarUI(); renderArvore(); }

  if (modo === "compose") {
    renderComposeCard();
    bootstrap(false).then(iniciarCompose).catch(() => {});
  } else {
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, detetarSelecao); } catch (e) {}
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, detetarSelecao); } catch (e) {}
    setInterval(detetarSelecao, 400);
    setInterval(() => carregarArvore().catch(() => {}), 25000);
    document.addEventListener("visibilitychange", () => { if (!document.hidden) { carregarArvore().catch(() => {}); detetarSelecao(); } });
    detetarSelecao();
    bootstrap(false);
  }
});

/* ---------- email selecionado ---------- */
function detetarSelecao() {
  const mbx = Office.context.mailbox;
  if (mbx && typeof mbx.getSelectedItemsAsync === "function") {
    mbx.getSelectedItemsAsync((res) => {
      if (!res || res.status !== Office.AsyncResultStatus.Succeeded) return detetarPorItem();
      const arr = res.value || [];
      const eraMulti = selecao.length > 1;
      if (arr.length > 1) {
        const sig = arr.map((x) => x.itemId).slice().sort().join("|");
        if (sig !== selecaoSig) {
          selecaoSig = sig;
          selecao = arr.map((x) => ({ itemId: x.itemId, subject: x.subject || "(sem assunto)" }));
          definirMulti();
        }
        return;
      }
      selecaoSig = "";
      selecao = arr.length ? [{ itemId: arr[0].itemId, subject: arr[0].subject || "(sem assunto)" }] : [];
      if (!arr.length) return definirEmail(null);
      const sel = arr[0];
      if (eraMulti || sel.itemId !== emailAtual.itemId) definirEmail({ itemId: sel.itemId, subject: sel.subject || "(sem assunto)" });
    });
  } else { detetarPorItem(); }
}
function detetarPorItem() {
  let item = null; try { item = Office.context.mailbox.item; } catch (e) {}
  const id = item ? item.itemId : null;
  if (id !== emailAtual.itemId) {
    definirEmail(item ? { itemId: id, subject: item.subject || "(sem assunto)",
      remetente: item.from ? (item.from.emailAddress || "").toLowerCase() : "",
      remetenteNome: item.from ? (item.from.displayName || item.from.emailAddress || "") : "" } : null);
  }
}
function definirEmail(d) {
  if (!d) { emailAtual = { itemId: null, subject: "", remetente: "", remetenteNome: "" }; renderEmail(); $("sugWrap").classList.add("hidden"); return; }
  emailAtual = { itemId: d.itemId, subject: d.subject || "", remetente: d.remetente || "", remetenteNome: d.remetenteNome || "" };
  renderEmail();
  if (emailAtual.remetente) calcularSugestao();
  else { $("sugWrap").classList.add("hidden"); obterRemetente(emailAtual.itemId); }
}
function definirMulti() {
  emailAtual = { itemId: null, subject: "", remetente: "", remetenteNome: "" };
  renderEmail();
  calcularSugestaoMulti();
}
async function obterRemetente(itemId) {
  try {
    const restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
    const r = await graphFetch("https://graph.microsoft.com/v1.0/me/messages/" + restId + "?$select=from");
    if (!r.ok) return;
    const m = await r.json();
    const addr = (m.from && m.from.emailAddress && m.from.emailAddress.address) || "";
    const nome = (m.from && m.from.emailAddress && m.from.emailAddress.name) || addr;
    if (itemId === emailAtual.itemId) { emailAtual.remetente = addr.toLowerCase(); emailAtual.remetenteNome = nome; renderEmail(); calcularSugestao(); }
  } catch (e) {}
}
function renderEmail() {
  const el = $("email");
  if (selecao.length > 1) {
    el.className = "card"; el.innerHTML = "";
    const a = document.createElement("div"); a.className = "assunto";
    a.textContent = "📥 " + selecao.length + " emails selecionados";
    el.appendChild(a);
    const max = 4;
    selecao.slice(0, max).forEach((s) => {
      const d = document.createElement("div"); d.className = "de"; d.textContent = "• " + s.subject; el.appendChild(d);
    });
    if (selecao.length > max) {
      const d = document.createElement("div"); d.className = "de"; d.textContent = "… e mais " + (selecao.length - max); el.appendChild(d);
    }
    return;
  }
  if (emailAtual.itemId && emailAtual.subject) {
    el.className = "card"; el.innerHTML = "";
    const a = document.createElement("div"); a.className = "assunto"; a.textContent = emailAtual.subject;
    const d = document.createElement("div"); d.className = "de"; d.textContent = emailAtual.remetenteNome ? ("de " + emailAtual.remetenteNome) : "";
    el.appendChild(a); el.appendChild(d);
  } else { el.className = "card vazio"; el.textContent = "Seleciona um email na lista."; }
}

/* ---------- modo compose (menu-ao-enviar) ---------- */
function renderComposeCard() {
  const el = $("email"); el.className = "card"; el.innerHTML = "";
  const a = document.createElement("div"); a.className = "assunto"; a.textContent = "📤 Arquivar uma cópia ao enviar";
  const d = document.createElement("div"); d.className = "de"; d.textContent = "Escolhe a pasta (ou pesquisa). Ao clicar, guardo a cópia e envio o email.";
  el.appendChild(a); el.appendChild(d);
}
function composeTo() {
  return new Promise((res) => {
    try { Office.context.mailbox.item.to.getAsync((r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded && r.value && r.value.length) res((r.value[0].emailAddress || "").toLowerCase());
      else res("");
    }); } catch (e) { res(""); }
  });
}
async function iniciarCompose() {
  const to = await composeTo();
  if (!to) return;
  emailAtual.remetente = to;
  const local = melhorDoUso(to);
  if (local && pastasById[local]) mostrarSugerida(local);
  else descobrirSugestao(to);
}
async function arquivarAoEnviar(folderId, folderName) {
  const item = Office.context.mailbox.item;
  status('A guardar cópia em "' + folderName + '" e a enviar…');
  try {
    const id = await new Promise((res) => { try { item.saveAsync((r) => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null)); } catch (e) { res(null); } });
    if (id) {
      const restId = Office.context.mailbox.convertToRestId(id, Office.MailboxEnums.RestVersion.v2_0);
      const r = await graphFetch("https://graph.microsoft.com/v1.0/me/messages/" + restId + "/copy", {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ destinationId: folderId }) });
      if (!r.ok) { status("⚠️ Não deu para guardar a cópia (" + r.status + "). Fecha e envia na mesma.", "err"); return; }
    }
    aprender(emailAtual.remetente, folderId);
    await new Promise((res) => { try { item.sessionData.setAsync("qf_done", "1", () => res()); } catch (e) { res(); } });
    status('✅ Cópia guardada em "' + folderName + '". A enviar…', "ok");
    try { item.sendAsync(() => {}); } catch (e) {}
  } catch (e) { status("⚠️ Erro: " + ((e && e.message) || e), "err"); }
}

/* ---------- auth ---------- */
async function initMsal() {
  if (!msalInstance) msalInstance = await msal.createNestablePublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: AUTHORITY }, cache: { cacheLocation: "localStorage" } });
}
async function ensureToken(allowPopup) {
  if (cachedToken) return cachedToken;
  await initMsal();
  try { cachedToken = (await msalInstance.acquireTokenSilent({ scopes: SCOPES })).accessToken; }
  catch (e) { if (!allowPopup) throw e; cachedToken = (await msalInstance.acquireTokenPopup({ scopes: SCOPES })).accessToken; }
  return cachedToken;
}
async function graphFetch(url, opts) {
  const build = (t) => Object.assign({}, opts, { headers: Object.assign({ Authorization: "Bearer " + t }, (opts && opts.headers) || {}) });
  let t = await ensureToken(true);
  let r = await fetch(url, build(t));
  if (r.status === 401) { cachedToken = null; t = await ensureToken(true); r = await fetch(url, build(t)); }
  return r;
}

/* ---------- carregar a arvore ---------- */
async function bootstrap(interactive) {
  try { await ensureToken(interactive); await carregarArvore(); mostrarUI(); if (emailAtual.remetente) calcularSugestao(); $("btnLogin").classList.add("hidden"); }
  catch (e) { if (!arvore.length) status("Liga a tua conta para começar."); $("btnLogin").classList.remove("hidden"); }
}
async function fetchFilhos(parentId) {
  const base = "https://graph.microsoft.com/v1.0/me/mailFolders";
  const url = (parentId ? base + "/" + parentId + "/childFolders" : base) + "?$top=250&$select=id,displayName,totalItemCount,childFolderCount";
  const r = await graphFetch(url);
  if (!r.ok) return [];
  return ((await r.json()).value || []);
}
async function construir(parentId, level) {
  const filhos = await fetchFilhos(parentId);
  filhos.sort((a, b) => (a.displayName || "").localeCompare(b.displayName || "", "pt"));
  await Promise.all(filhos.map(async (f) => {
    f.level = level; f.parentId = parentId; f.children = [];
    if (f.childFolderCount > 0 && level < 5) f.children = await construir(f.id, level + 1);
  }));
  return filhos;
}
async function carregarArvore() {
  const raiz = await construir(null, 0);
  arvore = raiz; pastasById = {}; indexar(arvore);
  localStorage.setItem(LS_PASTAS, JSON.stringify(arvore));
  renderArvore();
}
function indexar(lista) { (lista || []).forEach((p) => { pastasById[p.id] = p; if (p.children) indexar(p.children); }); }
function mostrarUI() { $("busca").classList.remove("hidden"); }

/* ---------- render ---------- */
function renderArvore() {
  const ul = $("lista"); ul.innerHTML = "";
  const q = $("busca").value.trim();
  if (q) return renderProcura(q, ul);
  arvore.forEach((p) => renderNo(ul, p));
}
function renderNo(ul, p) {
  ul.appendChild(itemPasta(p));
  if (expandido.has(p.id)) {
    (p.children || []).forEach((c) => renderNo(ul, c));
    ul.appendChild(itemNovaSubpasta(p));
  }
}
function itemPasta(p, sugerida, caminhoTxt) {
  const li = document.createElement("li");
  if (sugerida) li.className = "sug";
  li.style.paddingLeft = (8 + (caminhoTxt ? 0 : (p.level || 0) * 16)) + "px";
  const seta = document.createElement("span"); seta.className = "seta"; seta.textContent = "▶";
  if (!caminhoTxt) {
    if (expandido.has(p.id)) seta.classList.add("aberto");
    seta.addEventListener("click", (e) => { e.stopPropagation(); toggle(p.id); });
  } else { seta.className = "seta vazia"; }
  const nome = document.createElement("span"); nome.className = "nome";
  if (caminhoTxt) { nome.innerHTML = ""; const c = document.createElement("span"); c.className = "cam"; c.textContent = caminhoTxt + " › "; nome.appendChild(c); nome.appendChild(document.createTextNode(p.displayName || "")); }
  else nome.textContent = p.displayName || "(sem nome)";
  const cnt = document.createElement("span"); cnt.className = "cnt"; cnt.textContent = typeof p.totalItemCount === "number" ? String(p.totalItemCount) : "";
  const copy = document.createElement("span"); copy.className = "copy"; copy.title = "Copiar nome"; copy.textContent = "⧉";
  copy.addEventListener("click", (e) => { e.stopPropagation(); copiar(p.displayName); });
  li.appendChild(seta); li.appendChild(nome); li.appendChild(cnt); li.appendChild(copy);
  li.addEventListener("click", () => arquivar(p.id, p.displayName));
  return li;
}
function itemNovaSubpasta(parent) {
  const li = document.createElement("li"); li.className = "criar";
  li.style.paddingLeft = (8 + ((parent.level || 0) + 1) * 16) + "px";
  const n = document.createElement("span"); n.className = "nome"; n.textContent = "➕ Nova subpasta";
  li.appendChild(n);
  li.addEventListener("click", () => {
    li.innerHTML = "";
    const inp = document.createElement("input"); inp.className = "inp-sub"; inp.placeholder = "Nome + Enter";
    inp.addEventListener("click", (e) => e.stopPropagation());
    inp.addEventListener("keydown", (e) => { if (e.key === "Enter") criarSubpasta(parent, inp.value); else if (e.key === "Escape") renderArvore(); });
    li.appendChild(inp); inp.focus();
  });
  return li;
}
function caminho(p) {
  const parts = []; let cur = pastasById[p.parentId];
  while (cur) { parts.unshift(cur.displayName); cur = pastasById[cur.parentId]; }
  return parts.join(" › ");
}
function renderProcura(q, ul) {
  const ql = q.toLowerCase();
  Object.values(pastasById)
    .filter((p) => (p.displayName || "").toLowerCase().includes(ql))
    .sort((a, b) => (a.displayName || "").localeCompare(b.displayName || "", "pt"))
    .forEach((p) => ul.appendChild(itemPasta(p, false, caminho(p))));
  if (!Object.values(pastasById).some((p) => (p.displayName || "").toLowerCase() === ql)) {
    const li = document.createElement("li"); li.className = "criar";
    const n = document.createElement("span"); n.className = "nome"; n.textContent = '➕ Criar “' + q + '” (pasta principal)';
    li.appendChild(n); li.addEventListener("click", () => criarRaiz(q)); ul.appendChild(li);
  }
}
async function toggle(id) {
  if (expandido.has(id)) expandido.delete(id);
  else {
    expandido.add(id);
    const p = pastasById[id];
    if (p && p.childFolderCount > 0 && (!p.children || !p.children.length)) { p.children = await construir(id, (p.level || 0) + 1); indexar(p.children); }
  }
  renderArvore();
}

/* ---------- sugestao ---------- */
function calcularSugestao() {
  $("sugerido").innerHTML = ""; $("sugWrap").classList.add("hidden");
  const rem = emailAtual.remetente;
  if (!rem || !Object.keys(pastasById).length) return;
  const local = melhorDoUso(rem);
  if (local && pastasById[local]) return mostrarSugerida(local);
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
    const r = await graphFetch(url); if (!r.ok) return;
    const counts = {};
    ((await r.json()).value || []).forEach((m) => { const f = m.parentFolderId; if (pastasById[f] && !ehSistema(pastasById[f].displayName)) counts[f] = (counts[f] || 0) + 1; });
    let best = null, bestN = 0;
    Object.keys(counts).forEach((f) => { if (counts[f] > bestN) { best = f; bestN = counts[f]; } });
    if (best && emailAtual.remetente === rem) mostrarSugerida(best);
  } catch (e) {}
}
function mostrarSugerida(folderId) {
  const p = pastasById[folderId]; if (!p) return;
  $("sugerido").innerHTML = ""; $("sugerido").appendChild(itemPasta(p, true, caminho(p)));
  $("sugWrap").classList.remove("hidden");
}
async function obterRemetentesDe(itens) {
  const rems = []; const lote = 8;
  for (let i = 0; i < itens.length; i += lote) {
    const parte = await Promise.all(itens.slice(i, i + lote).map(async (it) => {
      try {
        const restId = Office.context.mailbox.convertToRestId(it.itemId, Office.MailboxEnums.RestVersion.v2_0);
        const r = await graphFetch("https://graph.microsoft.com/v1.0/me/messages/" + restId + "?$select=from");
        if (!r.ok) return null;
        const m = await r.json();
        return ((m.from && m.from.emailAddress && m.from.emailAddress.address) || "").toLowerCase();
      } catch (e) { return null; }
    }));
    parte.forEach((x) => { if (x) rems.push(x); });
  }
  return rems;
}
async function calcularSugestaoMulti() {
  $("sugerido").innerHTML = ""; $("sugWrap").classList.add("hidden");
  if (!Object.keys(pastasById).length) return;
  const sig = selecaoSig;
  const rems = await obterRemetentesDe(selecao.slice());
  if (selecaoSig !== sig || !rems.length) return;
  const agg = {};
  rems.forEach((rem) => {
    const m = uso[rem]; if (!m) return;
    Object.keys(m).forEach((fid) => { if (pastasById[fid] && !ehSistema(pastasById[fid].displayName)) agg[fid] = (agg[fid] || 0) + m[fid]; });
  });
  let best = null, bestN = 0;
  Object.keys(agg).forEach((fid) => { if (agg[fid] > bestN) { best = fid; bestN = agg[fid]; } });
  if (best) { if (selecaoSig === sig) mostrarSugerida(best); return; }
  const freq = {}; rems.forEach((r) => { freq[r] = (freq[r] || 0) + 1; });
  const remComum = Object.keys(freq).sort((a, b) => freq[b] - freq[a])[0];
  if (remComum) descobrirSugestaoMulti(remComum, sig);
}
async function descobrirSugestaoMulti(rem, sig) {
  try {
    const url = "https://graph.microsoft.com/v1.0/me/messages?$top=25&$select=parentFolderId&$search=" + encodeURIComponent('"from:' + rem + '"');
    const r = await graphFetch(url); if (!r.ok) return;
    const counts = {};
    ((await r.json()).value || []).forEach((m) => { const f = m.parentFolderId; if (pastasById[f] && !ehSistema(pastasById[f].displayName)) counts[f] = (counts[f] || 0) + 1; });
    let best = null, bestN = 0;
    Object.keys(counts).forEach((f) => { if (counts[f] > bestN) { best = f; bestN = counts[f]; } });
    if (best && selecaoSig === sig) mostrarSugerida(best);
  } catch (e) {}
}

/* ---------- arquivar ---------- */
async function arquivar(folderId, folderName) {
  if (modo === "compose") return arquivarAoEnviar(folderId, folderName);
  if (selecao.length > 1) return arquivarVarios(folderId, folderName);
  if (!emailAtual.itemId) { status("Seleciona primeiro um email.", "err"); return; }
  const rem = emailAtual.remetente, alvo = emailAtual.itemId;
  status('✅ Arquivado em "' + folderName + '".', "ok");
  aprender(rem, folderId);
  if (pastasById[folderId]) { pastasById[folderId].totalItemCount = (pastasById[folderId].totalItemCount || 0) + 1; renderArvore(); }
  try {
    const restId = Office.context.mailbox.convertToRestId(alvo, Office.MailboxEnums.RestVersion.v2_0);
    const r = await graphFetch("https://graph.microsoft.com/v1.0/me/messages/" + restId + "/move", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ destinationId: folderId }) });
    if (!r.ok) { if (pastasById[folderId]) { pastasById[folderId].totalItemCount = Math.max(0, (pastasById[folderId].totalItemCount || 1) - 1); renderArvore(); } status("⚠️ Afinal não deu para arquivar (" + r.status + ").", "err"); }
  } catch (e) { if (pastasById[folderId]) { pastasById[folderId].totalItemCount = Math.max(0, (pastasById[folderId].totalItemCount || 1) - 1); renderArvore(); } status("⚠️ Erro ao arquivar: " + ((e && e.message) || e), "err"); }
}
async function arquivarVarios(folderId, folderName) {
  const itens = selecao.slice();
  const n = itens.length;
  status("📤 A arquivar " + n + ' em "' + folderName + '"…');
  if (pastasById[folderId]) { pastasById[folderId].totalItemCount = (pastasById[folderId].totalItemCount || 0) + n; renderArvore(); }
  let ok = 0, falhou = 0;
  const lote = 8;
  for (let i = 0; i < itens.length; i += lote) {
    await Promise.all(itens.slice(i, i + lote).map(async (it) => {
      try {
        const restId = Office.context.mailbox.convertToRestId(it.itemId, Office.MailboxEnums.RestVersion.v2_0);
        const r = await graphFetch("https://graph.microsoft.com/v1.0/me/messages/" + restId + "/move", {
          method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ destinationId: folderId }) });
        if (r.ok) ok++; else falhou++;
      } catch (e) { falhou++; }
    }));
    if (i + lote < itens.length) status("📤 A arquivar… " + (ok + falhou) + "/" + n);
  }
  if (falhou && pastasById[folderId]) { pastasById[folderId].totalItemCount = Math.max(0, (pastasById[folderId].totalItemCount || falhou) - falhou); renderArvore(); }
  selecao = []; renderEmail();
  if (!falhou) status("✅ " + ok + ' emails arquivados em "' + folderName + '".', "ok");
  else if (ok) status("⚠️ " + ok + " arquivados, mas " + falhou + " falharam.", "err");
  else status("⚠️ Não deu para arquivar (" + falhou + " falharam).", "err");
}
function aprender(rem, folderId) {
  if (!rem) return;
  uso[rem] = uso[rem] || {}; uso[rem][folderId] = (uso[rem][folderId] || 0) + 1;
  localStorage.setItem(LS_USO, JSON.stringify(uso));
}

/* ---------- criar pastas ---------- */
async function criarSubpasta(parent, nome) {
  nome = (nome || "").trim(); if (!nome) return;
  status('A criar "' + nome + '" em "' + parent.displayName + '"…');
  try {
    const r = await graphFetch("https://graph.microsoft.com/v1.0/me/mailFolders/" + parent.id + "/childFolders", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ displayName: nome }) });
    if (!r.ok) { status("Não deu para criar (" + r.status + ").", "err"); return; }
    const nova = await r.json();
    nova.level = (parent.level || 0) + 1; nova.parentId = parent.id; nova.children = [];
    parent.children = parent.children || []; parent.children.push(nova);
    parent.children.sort((a, b) => (a.displayName || "").localeCompare(b.displayName || "", "pt"));
    parent.childFolderCount = (parent.childFolderCount || 0) + 1;
    pastasById[nova.id] = nova; expandido.add(parent.id);
    localStorage.setItem(LS_PASTAS, JSON.stringify(arvore));
    renderArvore();
    if (modo === "compose") await arquivarAoEnviar(nova.id, nova.displayName);
    else if (selecao.length > 1) await arquivarVarios(nova.id, nova.displayName);
    else if (emailAtual.itemId) await arquivar(nova.id, nova.displayName);
    else status('✅ Subpasta "' + nova.displayName + '" criada.', "ok");
  } catch (e) { status("Erro ao criar: " + ((e && e.message) || e), "err"); }
}
async function criarRaiz(nome) {
  nome = (nome || "").trim(); if (!nome) return;
  status('A criar a pasta "' + nome + '"…');
  try {
    const r = await graphFetch("https://graph.microsoft.com/v1.0/me/mailFolders", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ displayName: nome }) });
    if (!r.ok) { status("Não deu para criar (" + r.status + ").", "err"); return; }
    const nova = await r.json();
    nova.level = 0; nova.parentId = null; nova.children = [];
    arvore.push(nova); arvore.sort((a, b) => (a.displayName || "").localeCompare(b.displayName || "", "pt"));
    pastasById[nova.id] = nova; localStorage.setItem(LS_PASTAS, JSON.stringify(arvore));
    $("busca").value = ""; renderArvore();
    if (modo === "compose") await arquivarAoEnviar(nova.id, nova.displayName);
    else if (selecao.length > 1) await arquivarVarios(nova.id, nova.displayName);
    else if (emailAtual.itemId) await arquivar(nova.id, nova.displayName);
    else status('✅ Pasta "' + nova.displayName + '" criada.', "ok");
  } catch (e) { status("Erro ao criar: " + ((e && e.message) || e), "err"); }
}
function copiar(texto) {
  const ok = () => status('✅ Copiado: ' + texto, "ok");
  if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(texto).then(ok).catch(() => fallbackCopy(texto, ok));
  else fallbackCopy(texto, ok);
}
function fallbackCopy(texto, ok) {
  const ta = document.createElement("textarea"); ta.value = texto; ta.style.position = "fixed"; ta.style.opacity = "0";
  document.body.appendChild(ta); ta.focus(); ta.select();
  try { document.execCommand("copy"); ok(); } catch (e) { status("Não consegui copiar.", "err"); }
  document.body.removeChild(ta);
}
