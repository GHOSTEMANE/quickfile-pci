/* Arquivar — painel (Fase 2: validar a base; auth + pastas chegam na Fase 3) */
Office.onReady((info) => {
  const statusEl = document.getElementById("status");
  const contentEl = document.getElementById("content");

  if (info.host !== Office.HostType.Outlook) {
    statusEl.textContent = "Este add-in destina-se ao Outlook.";
    return;
  }

  statusEl.textContent = "Add-in a funcionar ✅";

  try {
    const item = Office.context.mailbox.item;
    if (item && item.subject) {
      contentEl.textContent = "Email selecionado: " + item.subject;
      contentEl.classList.remove("muted");
    } else {
      contentEl.textContent = "Seleciona um email para começar.";
    }
  } catch (e) {
    contentEl.textContent = "Pronto. (Nenhum email selecionado de momento.)";
  }
});
