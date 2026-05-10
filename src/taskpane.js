/* eslint-env browser */
/* global Office */

/**
 * Add-in Outlook « Envoyer vers Tentra ».
 * Logique :
 *  1. Au load Office.js, on extrait sujet / expéditeur / destinataires
 *     / date / body HTML / pièces jointes (en base64) du mail courant.
 *  2. On affiche un résumé dans le taskpane + active le bouton « Envoyer ».
 *  3. Au click, POST vers http://localhost:5180/api/import-email-staged.
 *  4. Tentra (s'il est ouvert) reçoit l'event SSE et ouvre la modale
 *     d'import. Sinon, le staging attend en disque jusqu'au prochain
 *     démarrage de Tentra.
 */

// URL de l'endpoint Tentra. Localhost par défaut. On peut la
// surcharger via un paramètre URL ?tentra=https://... pour des cas
// avancés (= tunneling, déploiement multi-machine), pas critique en V1.
const DEFAULT_TENTRA_ENDPOINT = 'http://localhost:5180/api/import-email-staged';
function resolveTentraEndpoint() {
  try {
    const params = new URLSearchParams(window.location.search);
    const override = params.get('tentra');
    if (override && /^https?:\/\//.test(override)) {
      return override.replace(/\/$/, '') + '/api/import-email-staged';
    }
  } catch { /* fallback default */ }
  return DEFAULT_TENTRA_ENDPOINT;
}
const TENTRA_ENDPOINT = resolveTentraEndpoint();

// Refs DOM. Réutilisés à la place de querySelector dans chaque handler.
const els = {
  status: null,
  statusText: null,
  details: null,
  subject: null,
  from: null,
  attachments: null,
  sendBtn: null,
  closeBtn: null,
};

/** Met à jour la zone de statut avec un type (info / success / error)
 *  et un message texte. Affiche/masque le spinner selon le type. */
function setStatus(type, message) {
  if (!els.status || !els.statusText) return;
  els.status.classList.remove('status-info', 'status-success', 'status-error');
  els.status.classList.add('status-' + type);
  els.statusText.textContent = message;
  const spinner = els.status.querySelector('.spinner');
  if (spinner) spinner.style.display = type === 'info' ? 'inline-block' : 'none';
}

/** Récupère les bytes d'une pièce jointe via Office.js et retourne
 *  une promesse de string base64. */
function getAttachmentBase64(attachmentId) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(
      attachmentId,
      { asyncContext: null },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          // result.value.format peut être 'base64' (= ce qu'on veut)
          // ou 'eml' / 'iCal' / 'url' selon le type d'attachement.
          // On ne traite que 'base64' en V1 ; les autres sont ignorés.
          if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
            resolve(result.value.content);
          } else {
            // Format non géré (= attachment item-as-msg, calendar event…)
            // On ignore plutôt que de planter.
            resolve(null);
          }
        } else {
          reject(result.error);
        }
      },
    );
  });
}

/** Récupère le body HTML du mail. Async. */
function getBodyHtml() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || '');
        } else {
          reject(result.error);
        }
      },
    );
  });
}

/** Construit le payload JSON à envoyer à Tentra à partir du mail
 *  courant. */
async function buildPayload() {
  const item = Office.context.mailbox.item;

  const subject = item.subject || '';
  const from = item.from
    ? { name: item.from.displayName || '', email: item.from.emailAddress || '' }
    : { name: '', email: '' };
  const to = (item.to || []).map(r => ({
    name: r.displayName || '', email: r.emailAddress || '',
  }));
  const cc = (item.cc || []).map(r => ({
    name: r.displayName || '', email: r.emailAddress || '',
  }));
  const date = item.dateTimeCreated
    ? new Date(item.dateTimeCreated).toISOString()
    : new Date().toISOString();
  const messageId = item.internetMessageId || undefined;

  const bodyHtml = await getBodyHtml();

  // Pièces jointes : on récupère le contenu binaire en base64 pour
  // chacune. On garde `contentId` pour pouvoir matcher plus tard
  // les images embedded (cid:) dans le bodyHtml côté pipeline Tentra.
  const rawAttachments = item.attachments || [];
  const attachments = [];
  for (const att of rawAttachments) {
    // Skip les attachements de type 'item' (= mail forward d'un autre
    // mail). Office.js permet de les traiter mais la conversion est
    // complexe — V1 on ignore.
    if (att.attachmentType === 'item') continue;
    try {
      const base64 = await getAttachmentBase64(att.id);
      if (!base64) continue;
      attachments.push({
        name: att.name || 'pièce jointe',
        contentType: att.contentType || 'application/octet-stream',
        contentId: att.contentId || undefined,
        base64,
      });
    } catch (err) {
      // eslint-disable-next-line no-console
      console.warn('[Tentra add-in] échec lecture pièce jointe', att.name, err);
    }
  }

  return {
    subject, from, to, cc, date, messageId, bodyHtml, attachments,
  };
}

/** Envoie le payload à Tentra. Retourne { ok, id } ou throw. */
async function sendToTentra(payload) {
  const resp = await fetch(TENTRA_ENDPOINT, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!resp.ok) {
    const errorText = await resp.text().catch(() => '');
    throw new Error(`HTTP ${resp.status} : ${errorText || 'unknown error'}`);
  }
  return resp.json();
}

/** Pré-affiche un résumé du mail dans le taskpane (= confirmation
 *  visuelle avant l'envoi). */
function renderPreview(payload) {
  els.subject.textContent = payload.subject || '(sans sujet)';
  const fromText = payload.from.name
    ? `${payload.from.name} <${payload.from.email}>`
    : (payload.from.email || '?');
  els.from.textContent = fromText;
  const attCount = payload.attachments.length;
  els.attachments.textContent = attCount === 0
    ? 'aucune'
    : attCount === 1
      ? '1 pièce jointe'
      : `${attCount} pièces jointes`;
  els.details.hidden = false;
}

/** Initialisation : récupère les refs DOM, déclenche l'extraction
 *  initiale, et wire les boutons. */
function init() {
  els.status = document.getElementById('status');
  els.statusText = els.status.querySelector('.status-text');
  els.details = document.getElementById('details');
  els.subject = document.getElementById('detail-subject');
  els.from = document.getElementById('detail-from');
  els.attachments = document.getElementById('detail-attachments');
  els.sendBtn = document.getElementById('send-btn');
  els.closeBtn = document.getElementById('close-btn');

  setStatus('info', 'Lecture du mail…');

  buildPayload()
    .then((payload) => {
      renderPreview(payload);
      setStatus('info', 'Prêt à envoyer.');
      els.sendBtn.disabled = false;
      els.sendBtn.addEventListener('click', () => {
        els.sendBtn.disabled = true;
        setStatus('info', 'Envoi vers Tentra…');
        sendToTentra(payload)
          .then(() => {
            setStatus('success', '✓ Envoyé. Validez l\'import dans Tentra.');
            els.closeBtn.hidden = false;
          })
          .catch((err) => {
            setStatus('error', `✗ Échec : ${err.message || err}`);
            els.sendBtn.disabled = false;
          });
      });
    })
    .catch((err) => {
      setStatus('error', `✗ Échec lecture mail : ${err.message || err}`);
    });

  els.closeBtn.addEventListener('click', () => {
    // Pas d'API Office.js pour fermer le taskpane proprement —
    // on déclenche un reload-blank pour que l'utilisateur voie
    // que l'opération est terminée et qu'il peut fermer manuellement.
    document.body.innerHTML = '<div class="closed">OK — tu peux fermer ce panneau.</div>';
  });
}

// Office.onReady est l'API moderne (= remplace Office.initialize).
// Garantit que l'API mailbox est prête avant qu'on l'utilise.
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    setStatus('error', 'Add-in chargé hors Outlook — abandon.');
    return;
  }
  init();
});
