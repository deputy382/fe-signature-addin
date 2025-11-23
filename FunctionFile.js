/*
 * FunctionFile.js — FE signature (standard placement)
 * - New compose: Outlook-managed bottom placement via setSignatureAsync
 * - Reply/Forward: insert just below the reply/forward header
 * - Works for ribbon ExecuteFunction (insertSignature) and LaunchEvent (onNewCompose)
 */

console.log('FunctionFile.js loaded');

var MARKER = 'FE_SIGNATURE_MARKER';

/* ========================= Helpers ========================= */

/** Wait until the body APIs return a string (up to maxMs). */
function waitForBodyReady(maxMs) {
  return new Promise(function (resolve) {
    var start = performance.now();
    (function check() {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded && typeof res.value === 'string') {
          return resolve();
        }
        if ((performance.now() - start) >= maxMs) return resolve(); // continue anyway
        setTimeout(check, 50);
      });
    })();
  });
}

/** Get compose type as a Promise ('NewMail' | 'Reply' | 'Forward'). */
function getComposeTypeAsync() {
  return new Promise(function (resolve) {
    if (Office.context?.mailbox?.item?.getComposeTypeAsync) {
      Office.context.mailbox.item.getComposeTypeAsync(function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          console.log('ComposeType:', res.value);
          resolve(res.value);
        } else {
          console.warn('ComposeType failed:', res.error);
          resolve('NewMail'); // safe default
        }
      });
    } else {
      resolve('NewMail');
    }
  });
}

/** Read body HTML. */
function getBodyHtmlAsync() {
  return new Promise(function (resolve, reject) {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          resolve(res.value || '');
        } else {
          console.error('getAsync(body HTML) failed:', res.error);
          reject(res.error);
        }
      }
    );
  });
}

/** Write body HTML. */
function setBodyHtmlAsync(newHtml) {
  return new Promise(function (resolve, reject) {
    Office.context.mailbox.item.body.setAsync(
      newHtml,
      { coercionType: Office.CoercionType.Html },
      function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          console.error('setAsync(body HTML) failed:', res.error);
          reject(res.error);
        }
      }
    );
  });
}

/** Insert signature just after common reply/forward header markers; else append at bottom. */
function insertAfterReplyHeader(html, sigHtml) {
  var patterns = [
    /<div[^>]*id=["']divRplyFwdMsg["'][^>]*>/i,         // Classic Outlook wrapper
    /<hr[^>]*>/i,                                       // Horizontal rule before quoted content
    /<div[^>]*>.*?On .*? wrote:\s*<\/div>/is,           // English textual header in a div
    /On .*? wrote:/i,                                   // English textual header (fallback)
    /<blockquote[^>]*>/i,                               // Quoted block (OWA/New Outlook)
    /<div[^>]*class=["'][^"']*(gmail_quote|moz-cite-prefix|yahoo_quoted|WordSection1)["'][^>]*>/i
  ];
  for (var i = 0; i < patterns.length; i++) {
    var match = html.match(patterns[i]);
    if (match) {
      var idx = html.indexOf(match[0]) + match[0].length;
      return html.slice(0, idx) + '\n' + sigHtml + '\n' + html.slice(idx);
    }
  }
  return html + '\n' + sigHtml; // fallback append
}

/* ========================= Signature builder ========================= */
function buildSignatureHtml(person) {
  var marker     = '<!-- ' + MARKER + ' -->\n';
  var openOuter  = '<div style="font-family:\'Segoe UI\', Arial, sans-serif; font-size:12px; line-height:1.35;">\n';
  var nameLine   = '<div style="font-size:13px; font-weight:600; color:#000;">' + person.displayName + '</div>\n';
  var titleLine  = '<div style="color:#000;">' + person.title + '</div>\n';
  var phoneLine  = '<div style="color:#000;">office: ' + person.officePhone +
                   (person.officeExt ? ' (' + person.officeExt + ')' : '') +
                   ' | cell: ' + person.mobile + '</div>\n';
  var emailLine  = '<div>mailto:' +
                   person.email + '</a></div>\n';
  var addrLine   = '<div style="color:#000;">' + person.address + ' | mailstop: ' + person.mailstop +
                   (person.site ? ' / ' + person.site : '') + '</div>\n';
  var closeOuter = '</div>';

  return (marker + openOuter + nameLine + titleLine + phoneLine + emailLine + addrLine + closeOuter).trim();
}

/* ========================= Shared implementation ========================= */
async function doInsertSignature() {
  // Make sure body APIs are ready
  await waitForBodyReady(400);

  // Optional: disable the client-managed signature to prevent duplicates
  if (Office.context?.mailbox?.item?.disableClientSignatureAsync) {
    Office.context.mailbox.item.disableClientSignatureAsync(function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn('disableClientSignatureAsync failed:', res.error);
      }
    });
  }

  var composeType = await getComposeTypeAsync();

  // Static person block for v1; we’ll swap to Graph later
  var person = {
    displayName: 'Shane Francis',
    title: 'Systems Administrator B IV',
    officePhone: '3303238382',
    officeExt: '850-2601',
    mobile: '330-323-8382',
    email: 'sfrancis@firstenergycorp.com',
    address: '341 White Pond Drive, Akron, OH 44320',
    mailstop: 'A-FEHQ-A2',
    site: 'Akron FirstEnergy Headquarters'
  };

  var sigHtml = buildSignatureHtml(person);

  if (composeType === 'NewMail') {
    // Let Outlook place the signature at the STANDARD BOTTOM for new mail
    return new Promise(function (resolve) {
      Office.context.mailbox.item.body.setSignatureAsync(
        sigHtml,
        { coercionType: Office.CoercionType.Html }, // HTML signature
        function (res) {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            console.warn('setSignatureAsync failed; falling back to setBody:', res.error);
            // Fallback: append to bottom
            getBodyHtmlAsync()
              .then(function (bodyHtml) { return setBodyHtmlAsync(bodyHtml + '\n' + sigHtml); })
              .then(function () { console.log('Signature inserted (fallback bottom).'); resolve(); })
              .catch(function (err) { console.error('Fallback setBody failed:', err); resolve(); });
          } else {
            console.log('Signature inserted (Outlook-managed bottom).');
            resolve();
          }
        }
      );
    });
  } else {
    // Reply or Forward: place directly under the reply header
    var bodyHtml = await getBodyHtmlAsync();

    // Idempotence: skip if already present
    var hasMarker = bodyHtml.indexOf('<!-- ' + MARKER + ' -->') !== -1 || bodyHtml.indexOf(MARKER) !== -1;
    if (hasMarker) {
      console.log('Signature already present; skipping.');
      return;
    }

    var newHtml = insertAfterReplyHeader(bodyHtml, sigHtml);
    await setBodyHtmlAsync(newHtml);
    console.log('Signature inserted (below reply header).');
  }
}

/* ========================= Command + Event wrappers ========================= */

// Ribbon button: ExecuteFunction -> insertSignature
async function insertSignature(event) {
  try { await doInsertSignature(); }
  catch (err) { console.error('❌ Button insertSignature failed:', err); }
  finally { if (event && typeof event.completed === 'function') { event.completed(); } }
}

// Event-based autorun: LaunchEvent -> onNewCompose
async function onNewCompose(event) {
  try { await doInsertSignature(); }
  catch (err) { console.error('❌ Autorun onNewCompose failed:', err); }
  finally { if (event && typeof event.completed === 'function') { event.completed(); } }
}

/* Expose for the ribbon button (ExecuteFunction). */
window.insertSignature = insertSignature;

/* ========================= Associate event handler ========================= */
try {
  // Bind the event handler name used in the manifest
  Office.actions.associate('onNewCompose', onNewCompose);
} catch (e) {
  console.debug('Initial associate deferred:', e);
}

Office.onReady(function () {
  console.log('Autorun runtime loaded:', Office.context.platform);

  // Bind again once Office is ready (safe to double-associate)
  try { Office.actions.associate('onNewCompose', onNewCompose); }
  catch (e) { console.debug('Associate onReady already bound:', e); }
