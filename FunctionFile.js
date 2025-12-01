
/* FunctionFile.js — FE Signature (commands + autorun) */
console.log('FunctionFile.js loaded');

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Office.js is ready for Outlook");

    // Expose handlers globally for ribbon and autorun
    window.insertSignature = insertSignature;
    window.onNewCompose = onNewCompose;
  }
});

// ---- Constants
const SIG_MARKER = 'FE_SIGNATURE_MARKER';
const SIG_COMMENT = `<!-- ${SIG_MARKER} -->`;

// ---- Utilities
function waitForBodyReady(maxMs = 400) {
  return new Promise((resolve) => {
    const start = performance.now();
    (function check() {
      try {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (res) => {
          const ok = res.status === Office.AsyncResultStatus.Succeeded && typeof res.value === 'string';
          if (ok) return resolve();
          if ((performance.now() - start) >= maxMs) return resolve(); // proceed anyway
          setTimeout(check, 50);
        });
      } catch {
        resolve();
      }
    })();
  });
}

function getComposeTypeAsync() {
  return new Promise((resolve) => {
    const fn = Office.context?.mailbox?.item?.getComposeTypeAsync;
    if (!fn) return resolve('newMail');
    fn((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded && res.value?.composeType) {
        console.log('ComposeType:', res.value.composeType);
        resolve(res.value.composeType); // expected: newMail, reply, forward
      } else {
        console.warn('getComposeTypeAsync failed:', res.error);
        resolve('newMail');
      }
    });
  });
}

function getBodyHtmlAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve(res.value || '');
      } else {
        reject(res.error);
      }
    });
  });
}

function setBodyHtmlAsync(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

function insertBelowReplyHeader(html, sigHtml) {
  const patterns = [
    /<div[^>]*id=["']divRplyFwdMsg["'][^>]*>/i,
    /<hr[^>]*>/i,
    /<blockquote[^>]*>/i,
    /<div[^>]*class=["'][^"']*(gmail_quote|moz-cite-prefix|yahoo_quoted|WordSection1)["'][^>]*>/i,
    /On .+ wrote:/i
  ];
  for (const re of patterns) {
    const m = html.match(re);
    if (m) {
      const idx = html.indexOf(m[0]) + m[0].length;
      return html.slice(0, idx) + '\n' + sigHtml + '\n' + html.slice(idx);
    }
  }
  return html + '\n' + sigHtml;
}

function buildSignatureHtml() {
  const lines = [
    '<div style="font-family:\'Segoe UI\', Arial, sans-serif; font-size:12px; line-height:1.35;">',
    '<div style="font-size:13px; font-weight:600; color:#000;">Shane Francis</div>',
    '<div style="color:#000;">Systems Administrator B IV</div>',
    '<div style="color:#000;">office: 850-2601 cell: 330-323-8382</div>',
    '<div>mailto:sfrancis@firstenergycorp.com</div>',
    '<div style="color:#000;">341 White Pond Drive, Akron, OH 44320 mailstop: A-FEHQ-A2 / Akron FirstEnergy HQ</div>',
    '</div>'
  ];
  return (SIG_COMMENT + '\n' + lines.join('\n')).trim();
}

async function doInsertSignature() {
  await waitForBodyReady(400);
  if (Office.context?.mailbox?.item?.disableClientSignatureAsync) {
    Office.context.mailbox.item.disableClientSignatureAsync((res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn('disableClientSignatureAsync failed:', res.error);
      }
    });
  }
  const composeType = await getComposeTypeAsync();
  const sigHtml = buildSignatureHtml();
  if (composeType === 'newMail') {
    return new Promise((resolve) => {
      try {
        Office.context.mailbox.item.setSignatureAsync(sigHtml, (res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            console.warn('setSignatureAsync failed; falling back to append:', res.error);
            getBodyHtmlAsync()
              .then((bodyHtml) => setBodyHtmlAsync(bodyHtml + '\n' + sigHtml))
              .then(() => { console.log('Signature inserted (fallback bottom).'); resolve(); })
              .catch((err) => { console.error('Fallback setBody failed:', err); resolve(); });
          } else {
            console.log('Signature inserted (Outlook-managed bottom).');
            resolve();
          }
        });
      } catch (e) {
        console.warn('setSignatureAsync threw; attempting fallback:', e);
        getBodyHtmlAsync()
          .then((bodyHtml) => setBodyHtmlAsync(bodyHtml + '\n' + sigHtml))
          .then(() => console.log('Signature inserted (fallback bottom).'))
          .catch((err) => console.error('Fallback setBody failed:', err));
      }
    });
  } else {
    const bodyHtml = await getBodyHtmlAsync();
    if (bodyHtml.includes(SIG_MARKER) || bodyHtml.includes(SIG_COMMENT)) {
      console.log('Signature already present; skipping.');
      return;
    }
    const newHtml = insertBelowReplyHeader(bodyHtml, sigHtml);
    await setBodyHtmlAsync(newHtml);
    console.log('Signature inserted (below reply header).');
  }
}

// ---- Ribbon button handler
async function insertSignature(event) {
  try { await doInsertSignature(); }
  catch (err) { console.error('❌ insertSignature failed:', err); }
  finally { if (event && typeof event.completed === 'function') event.completed(); }
}

async function onNewCompose(event) {
  console.log('Autorun works!');
  try { await doInsertSignature(); }
  catch (err) { console.error('❌ onNewCompose failed:', err); }
  finally { if (event && typeof event.completed === 'function') event.completed(); }
}
