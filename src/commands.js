/* Menu-ao-enviar (OnMessageSend): ao enviar, mostra um aviso cujo botao abre
   o painel Arquivar PCI (o mesmo menu completo) para escolher a pasta.
   O painel (app.js, modo compose) trata de guardar a copia e enviar. */

function lerSessao(item, k) {
  return new Promise((res) => {
    try { item.sessionData.getAsync(k, (r) => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null)); }
    catch (e) { res(null); }
  });
}

async function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;
  try {
    // Se o painel ja tratou este envio (guardou a copia e mandou enviar), deixar seguir.
    const done = await lerSessao(item, "qf_done");
    if (done) { event.completed({ allowEvent: true }); return; }
    // Caso contrario, mostrar o aviso com o botao que abre o painel completo.
    event.completed({
      allowEvent: false,
      errorMessage: "Arquivar uma cópia deste email? Toca em «Escolher pasta» para abrir o menu e escolher onde guardar — ou «Enviar na mesma» para enviar sem guardar.",
      cancelLabel: "Escolher pasta",
      commandId: "ArquivarComposeButton"
    });
  } catch (e) {
    try { event.completed({ allowEvent: true }); } catch (e2) {}
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
