
/*
 * FunctionFile.js — FE signature (standard placement), button + autorun
 * - New compose: append at bottom
 * - Reply/Forward: insert just below the reply/forward header
 * - Prevents duplicates via marker
 * - Works for both ribbon ExecuteFunction and OnNewMessageCompose autorun
 */

console.log('FunctionFile.js loaded');

var MARKER = 'FE_SIGNATURE_MARKER';

/* ========================= Helpers ========================= */

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

/* ========================= Shared handler for button + autorun ========================= */
async function insertSignature(event) {
  try {
    // If the event parameter is missing (button click), create a stub so we can always call completed().
    var evt = event || { completed: function () {} };

    // Small defer helps on some builds where body isn’t immediately ready for new compose
    await new Promise(function (r) { setTimeout(r, 25); });

    var bodyHtml = await getBodyHtmlAsync();

    // Idempotence: skip if already present
    var hasMarker = bodyHtml.indexOf('<!-- ' + MARKER + ' -->') !== -1 || bodyHtml.indexOf(MARKER) !== -1;
    if (hasMarker) {
      console.log('Signature already present; skipping.');
      return;
    }

    // v1 static user block (we’ll replace with Graph later)
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
    var newHtml = insertAfterReplyHeader(bodyHtml, sigHtml);
    await setBodyHtmlAsync(newHtml);

    console.log('Signature inserted.');
  } catch (err) {
    console.error('❌ Signature insertion failed:', err);
  } finally {
    if (event && typeof event.completed === 'function') {
      event.completed();
    }
  }
}

/* Make the function AVAILABLE to the ribbon button (ExecuteFunction). */
window.insertSignature = insertSignature;

/* ========================= Early association to catch autorun before onReady ========================= */
try {
  // Bind as soon as the script loads (helps if the sandbox fires before onReady)
  Office.actions.associate('insertSignature', insertSignature);
} catch (e) {
  // on some clients Office may not be ready yet—onReady below will bind again
  console.debug('Initial associate deferred:', e);
}

/* ========================= Consolidated Office.onReady ========================= */
Office.onReady(function () {
  console.log('Autorun runtime loaded:', Office.context.platform);

  // Bind again once Office is ready (safe to double-associate)
  try {
    Office.actions.associate('insertSignature', insertSignature);
  } catch (e) {
    console.debug('Associate onReady already bound:', e);
  }

  // Optional diagnostics: what compose type are we in? (NewMail | Reply | Forward)
  if (Office.context && Office.context.mailbox && Office.context.mailbox.item &&
      typeof Office.context.mailbox.item.getComposeTypeAsync === 'function') {
    Office.context.mailbox.item.getComposeTypeAsync(function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        console.log('ComposeType:', res.value);
      } else {
        console.warn('ComposeType failed:', res.error);
      }
    });
  }
});
