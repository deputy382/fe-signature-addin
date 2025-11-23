
/*
 * FunctionFile.js — FE signature (standard placement), button + autorun fallback
 * - New compose: append at bottom
 * - Reply/Forward: insert just below the reply/forward header
 * - Prevents duplicates via marker
 * - Works for ribbon ExecuteFunction and OnNewMessageCompose autorun
 * - Adds a tiny deferred fallback call so compose always gets the signature
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
    var evt = event || { completed: function () {} };

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

/* Expose for the ribbon button (ExecuteFunction). */
window.insertSignature = insertSignature;

/* ========================= Early association + autorun fallback shim ========================= */
try {
  // Bind as soon as the script loads (helps if the sandbox triggers before onReady)
  Office.actions.associate('insertSignature', insertSignature);
} catch (e) {
  console.debug('Initial associate deferred:', e);
}

Office.onReady(function () {
  console.log('Autorun runtime loaded:', Office.context.platform);

  // Bind again once Office is ready (safe to double-associate)
  try {
    Office.actions.associate('insertSignature', insertSignature);
  } catch (e) {
    console.debug('Associate onReady already bound:', e);
  }

  // Detect compose type and run a tiny deferred fallback
  if (Office.context && Office.context.mailbox && Office.context.mailbox.item &&
      typeof Office.context.mailbox.item.getComposeTypeAsync === 'function') {
    Office.context.mailbox.item.getComposeTypeAsync(function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        console.log('ComposeType:', res.value);
        // Fallback shim: if the event didn’t fire, run the handler after a short delay.
        setTimeout(function () {
          // The marker prevents double insert if the event already fired.
          try { window.insertSignature(); } catch (ex) { console.warn('Fallback insertSignature failed:', ex); }
        }, 50);
      } else {
        console.warn('ComposeType failed:', res.error);
      }
    });
  }
});
