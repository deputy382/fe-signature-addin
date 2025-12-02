/* FunctionFileManual.js — FE Signature (commands + autorun) */
console.log('FunctionFileManual.js loaded');

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
        resolve(res.value.composeType);
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

// ---- Graph helpers
async function getGraphToken() {
  try {
    // Prompts for consent if needed; requires WebApplicationInfo in manifest
    const token = await OfficeRuntime.auth.getAccessToken({ allowConsentPrompt: true });
    return token;
  } catch (e) {
    console.warn('getAccessToken failed:', e);
    return null;
  }
}

async function getUserProfileFromGraph() {
  const token = await getGraphToken();
  if (!token) return null;
  try {
    const res = await fetch(
      'https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mail,mobilePhone,officeLocation,businessPhones',
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Graph ${res.status}: ${text}`);
    }
    return await res.json();
  } catch (e) {
    console.warn('Graph call failed:', e);
    return null;
  }
}

function escapeHtml(s) {
  if (!s) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ---- Signature builder (Graph-powered with safe fallback)
async function buildSignatureHtml() {
  let p = await getUserProfileFromGraph();

  // Fallbacks if Graph is unavailable
  const fallback = {
    displayName: 'Shane Francis',
    jobTitle: 'Systems Administrator B IV',
    mail: 'sfrancis@firstenergycorp.com',
    officeLocation: '341 White Pond Drive, Akron, OH 44320 &nbsp; mailstop: A-FEHQ-A2 / Akron FirstEnergy HQ',
    businessPhones: ['850-2601'],
    mobilePhone: '330-323-8382'
  };

  const name = escapeHtml(p?.displayName || fallback.displayName);
  const title = escapeHtml(p?.jobTitle || fallback.jobTitle);
  const email = escapeHtml(p?.mail || fallback.mail);
  const phones = Array.isArray(p?.businessPhones) ? p.businessPhones.filter(Boolean) : [];
  const officePhone = escapeHtml(phones[0] || fallback.businessPhones[0]);
  const mobile = escapeHtml(p?.mobilePhone || fallback.mobilePhone);
  const officeLoc = (p?.officeLocation ? escapeHtml(p.officeLocation) : fallback.officeLocation);

  const phoneLine = (officePhone || mobile)
    ? `office: ${officePhone}${mobile ? ' &nbsp;&nbsp; cell: ' + mobile : ''}` : '';

  const sig = `
    <div style="font-family:'Segoe UI', Arial, sans-serif; font-size:12px; line-height:1.35;">
      <div style="font-size:13px; font-weight:600; color:#000;">${name}</div>
      ${title ? `<div style="color:#000;">${title}</div>` : ''}
      ${phoneLine ? `<div style="color:#000;">${phoneLine}</div>` : ''}
      ${email ? `<div>mailto:${email}${email}</a></div>` : ''}
      ${officeLoc ? `<div style="color:#000;">${officeLoc}</div>` : ''}
    </div>
  `.trim();

  return (SIG_COMMENT + '\n' + sig).trim();
}

// ---- Core implementation
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
  const sigHtml = await buildSignatureHtml(); // <-- now async

  if (composeType === 'newMail') {
    return new Promise((resolve) => {
      try {
        // Using body.setSignatureAsync per your existing manual implementation
        Office.context.mailbox.item.body.setSignatureAsync(sigHtml, (res) => {
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

// ---- Wrappers (required for commands/events)
async function insertSignature(event) {
  try { await doInsertSignature(); }
  catch (err) { console.error('❌ insertSignature failed:', err); }
  finally { if (event && typeof event.completed === 'function') event.completed(); }
}

async function onNewCompose(event) {
  try { await doInsertSignature(); }
  catch (err) { console.error('❌ onNewCompose failed:', err); }
  finally { if (event && typeof event.completed === 'function') event.completed(); }
}

// ---- Expose and associate only after Office is ready
Office.onReady(() => {
  console.log('Autorun runtime loaded:', Office.context.platform);
  window.insertSignature = insertSignature;
  window.onNewCompose = onNewCompose;
  try { Office.actions.associate('onNewCompose', onNewCompose); } catch (e) { /* already bound */ }
});

// Optional legacy no-op
Office.initialize = () => {};
