# Add-in Outlook de Arquivamento à Medida — Plano de Implementação

> **Para quem executa:** usar `superpowers:executing-plans` (execução inline com checkpoints). Projeto com forte integração externa (Azure, M365 admin, Outlook) e ações do administrador — não adequado a subagentes isolados. Steps com checkbox `- [ ]`.

**Goal:** Recriar a experiência do QuickFile do pai do Israel — arquivar recebidos com sugestão de pasta + menu automático ao enviar (Send&File) — como add-in web Outlook à medida, gratuito e permanente, na conta Arenor.

**Architecture:** Add-in web Outlook (manifest add-in-only, no estilo do QuickFile365). Task pane em HTML/JS estático. Office.js para a integração (botão de leitura, botão de composição, evento `OnMessageSend`). Microsoft Graph para listar pastas, mover emails e calcular sugestões por histórico. Autenticação por **Nested App Authentication (NAA)** via MSAL.js — tokens Graph no cliente, **sem backend**. Site estático HTTPS gratuito. **Deployment central pelo admin (Israel)** no M365 Admin Center — é o que ativa o `OnMessageSend` e dá admin consent às permissões.

**Tech Stack:** HTML/CSS/JS (vanilla), Office.js, MSAL.js (NAA), Microsoft Graph REST, manifest XML add-in-only. Hosting: Azure Static Web Apps (grátis; Israel é admin Azure) ou GitHub Pages.

## Global Constraints
- Conta-alvo: M365 Arenor (`israelsalvadorjesus@arenor.es`), Outlook for Mac 16.110 (New Outlook ativo), macOS Tahoe 26.5.
- Sem subscrição e sem backend pago: tudo estático + hosting grátis.
- Sugestão de pasta = por **histórico** (remetente/domínio → pastas onde já arquivou). Sem LLM (custo zero). Inteligência de contexto fica para evolução futura.
- `SendMode` do menu-ao-enviar = `PromptUser` (nunca bloqueia o envio; degrada para "enviar à mesma" se algo falhar).
- Deployment **central via admin** (não sideload) — necessário para o `OnMessageSend` disparar no novo Outlook.

## Riscos / validação antecipada (SPIKE primeiro)
A peça mais arriscada é **autenticação Graph (NAA) + deployment central**. As Fases 1–3 são um spike: se o add-in autenticar e **listar as pastas** da conta Arenor dentro do Outlook, o caminho está validado e o resto é incremental. As Fases 4–8 são refinadas (código fino) **depois** do spike passar — não vale a pena escrever código especulativo antes disso.

---

### Fase 1: App Registration no Azure (ação do Israel, guiada clique-a-clique)
**Entrega:** uma identidade de aplicação no Entra ID (Azure AD) da Arenor, com as permissões certas e admin consent dado. Output: **Client ID** (+ Tenant ID).
- [ ] Criar App Registration (Entra admin center → App registrations → New).
- [ ] Plataforma **Single-page application (SPA)**; adicionar redirect URI do broker NAA (`brk-multihub://<dominio-do-hosting>`) + a URL do hosting.
- [ ] API permissions (Microsoft Graph, **delegated**): `Mail.ReadWrite`, `User.Read`, `MailboxFolder.ReadWrite` (se aplicável) → **Grant admin consent** (Israel é admin).
- [ ] Expor a app para NAA conforme guia MSAL; copiar **Application (client) ID** e **Directory (tenant) ID**.
- **Checkpoint:** Client ID e Tenant ID em mãos.

### Fase 2: Esqueleto do add-in + hosting + aparecer no Outlook
**Entrega:** um add-in mínimo (botão na leitura que abre um task pane "olá") visível no Outlook do Israel.
- [ ] Estrutura do projeto: `src/taskpane.html`, `src/taskpane.js`, `src/commands.html`, `manifest.xml`, ícones.
- [ ] Manifest add-in-only mínimo (baseado no padrão `~/Downloads/QuickFile365_SmartAlert.xml`): `MessageReadCommandSurface` + task pane; URLs a apontar para o hosting.
- [ ] Publicar o site estático (Azure Static Web Apps / GitHub Pages) → obter URL HTTPS.
- [ ] Instalar (sideload p/ teste rápido) e confirmar o botão no Outlook.
- **Checkpoint:** botão aparece e abre o task pane.

### Fase 3 (SPIKE CRÍTICO): Auth NAA + listar pastas
**Entrega:** o task pane mostra as pastas reais da conta Arenor.
- [ ] Integrar MSAL.js com config NAA (`createNestablePublicClientApplication`), usar Client/Tenant ID da Fase 1.
- [ ] Obter token Graph silencioso; `GET https://graph.microsoft.com/v1.0/me/mailFolders?$top=100`.
- [ ] Renderizar a árvore de pastas no task pane.
- **Checkpoint DECISIVO:** se as pastas aparecerem, a arquitetura está validada. Se NAA falhar no Outlook Mac, recuar para SSO + Azure Function (OBO) — ainda grátis (consumption).

### Fase 4: Arquivar email recebido para uma pasta
**Entrega:** botão "Arquivar aqui" move o email aberto para a pasta escolhida.
- [ ] Ler o item atual (`Office.context.mailbox.item.itemId`); converter para REST id se necessário.
- [ ] `POST /me/messages/{id}/move` com `destinationId`.
- [ ] Feedback no task pane + caixa de pesquisa de pastas (filtro client-side).
- **Checkpoint:** email sai da inbox e aparece na pasta (confirmar no Outlook e no servidor).

### Fase 5: Sugestão por histórico (TDD na lógica pura)
**Entrega:** ao abrir um email, as pastas prováveis aparecem no topo.
- [ ] **TDD:** função pura `suggestFolders(senderHistory) -> [folderId]` (ordena por frequência/recência). Escrever teste que falha → implementar → passar.
- [ ] Recolha via Graph: `GET /me/messages?$search="from:<email>"&$select=parentFolderId` → contar por `parentFolderId` → top N (excluir Inbox/Sent).
- [ ] Cache simples remetente→pastas (localStorage) para rapidez.
- **Checkpoint:** sugestão acerta na maioria dos remetentes recorrentes.

### Fase 6: Criar pasta nova
**Entrega:** botão "Nova pasta" cria e usa logo a pasta.
- [ ] `POST /me/mailFolders` (ou subpasta: `/me/mailFolders/{id}/childFolders`).
- [ ] Atualizar a lista e selecionar a nova pasta.
- **Checkpoint:** pasta criada aparece no Outlook.

### Fase 7: Send&File (OnMessageSend / Smart Alert)
**Entrega:** ao clicar Enviar, salta o menu a perguntar (enviar só / guardar em pasta X / nova pasta).
- [ ] Manifest: `LaunchEvent Type="OnMessageSend" SendMode="PromptUser"` + runtime do evento (como no manifest de referência).
- [ ] Handler de envio: abre o task pane para escolher pasta (ou "enviar à mesma"); `event.completed`.
- [ ] Após envio, mover o item de **Sent Items** para a pasta escolhida (fila de processamento, como o QuickFile365 faz, para fiabilidade).
- **Checkpoint:** prompt aparece ao enviar; enviado é arquivado na pasta certa.

### Fase 8: Deployment central + admin consent + teste E2E
**Entrega:** add-in instalado para a conta via admin; tudo a funcionar de forma estável.
- [ ] M365 Admin Center → Settings → Integrated Apps → **Deploy add-in** (Upload custom apps → manifest) → atribuir ao Israel (ou grupo).
- [ ] Confirmar admin consent das permissões Graph.
- [ ] Teste end-to-end: arquivar recebido (com sugestão) + criar pasta + enviar email e ver o menu-ao-enviar a arquivar.
- **Checkpoint:** experiência equivalente à do QuickFile do pai, na conta Arenor.

---

## Notas
- O nome do add-in no Outlook fica a decidir (ex.: "Arquivar" / "PCIFile").
- "Todas as contas": limitação da Microsoft (add-ins só na conta primária no Mac) — fora do âmbito deste MVP; reavaliar depois.
- Manutenção: o add-in parte se a Microsoft mudar APIs; o hosting tem de se manter online. Aceite como custo de ser à medida e grátis.
