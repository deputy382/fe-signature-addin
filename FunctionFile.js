
/*
 * FunctionFile.js — Event-based signature insertion in the standard places:
 * - New compose: append at bottom of the body.
 * - Reply/Forward: insert just under the reply/forward header; else append at bottom.
 */

console.log("FunctionFile.js loaded");

// Unique marker so we don't insert twice
const MARKER = "FE_SIGNATURE_MARKER";

// Build your signature HTML (unescaped real HTML)
function buildSignatureHtml() {
  return `
    <!-- ${MARKER} -->
    <table style="font-family:'Segoe UI', Arial, sans-serif; font-size:12px; line-height:1.35;">
      <tr>
        <td style="padding:6px 0;">
          <strong>FirstEnergy</strong><br/>
          Employee Name | Title<br/>
          Department<br/>
          https://www.firstenergycorp.com
            www.firstenergycorp.com
          </a><br/>
          <span>Email: user@firstenergycorp.com</span>
        </td>
      </tr>
    </table>
  `.trim();
}

// Heuristics: find a likely reply/forward header and insert *after* it.
// If no header is found, return html with signature appended at bottom.
function insertAfterReplyHeader(html, sigHtml) {
  // Common markers seen across Outlook clients (Classic/OWA/New Mac/New Windows)
  const patterns = [
    // Classic Outlook: div that wraps quoted message when replying/forwarding
    /<div[^>]*id=["']divRplyFwdMsg["'][^>]*>/i,

    // Horizontal rule inserted before quoted content
    /<hr[^>]*>/i,

    // Generic textual header (English) — adapt if you localize
    /<div[^>]*>.*?On .*? wrote:\s*<\/div>/is,
    /On .*? wrote:/i,

    // Quoted content container blocks often used by OWA/New Outlook
    /<blockquote[^>]*>/i,
    /<div[^>]*class=["'][^"']*(gmail_quote|moz-cite-prefix|yahoo_quoted|WordSection1)["'][^>]*>/i
  ];

  for (const pattern of patterns) {
    const match = html.match(pattern);
    if (match) {
      const idx = html.indexOf(match[0]) + match[0].length;
      // Insert right after the marker
      return html.slice(0, idx) + "\n" + sigHtml + "\n" + html.slice(idx);
    }
  }

  // Fallback: append at bottom
  return `${html}\n${sigHtml}`;
}

// Read body (HTML)
function getBodyHtmlAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          resolve(res.value || "");
        } else {
          reject(res.error);
        }
      }
    );
  });
}

// Write body (HTML)
function setBodyHtmlAsync(newHtml) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(
      newHtml,
      { coercionType: Office.CoercionType.Html },
      (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(res.error);
      }
    );
  });
}

// Prepend signature at top (not used for “standard” placement but kept as utility)
function prependSignatureAsync(sigHtml) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.prependAsync(
      sigHtml,
      { coercionType: Office.CoercionType.Html },
      (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(res.error);
      }
    );
  });
}

// Decide placement: reply/forward vs. new compose.
// We use body inspection for header markers; if none, treat as new (append).
async function insertSignatureStandardPlacement() {
  const bodyHtml = await getBodyHtmlAsync();

  // Idempotence: skip if marker already present
  const hasMarker =
    bodyHtml.includes(`<!-- ${MARKER} -->`) ||
    bodyHtml.includes(MARKER);
  if (hasMarker) {
    console.log("Signature already present; skipping.");
    return;
  }

  const sigHtml = buildSignatureHtml();

  // If we detect a reply/forward header, insert under it; else append bottom.
  const newHtml = insertAfterReplyHeader(bodyHtml, sigHtml);
  if (newHtml !== bodyHtml) {
    await setBodyHtmlAsync(newHtml);
    console.log("Signature inserted at standard location.");
  } else {
    // As a fallback (shouldn’t happen), append signature at bottom
    const appendedHtml = `${bodyHtml}\n${sigHtml}`;
    await setBodyHtmlAsync(appendedHtml);
    console.log("Signature appended at bottom (fallback).");
  }
}

// Event-based handler called by LaunchEvent OnNewMessageCompose
async function insertSignature(event) {
  try {
    await insertSignatureStandardPlacement();
  } catch (err) {
    console.error("Signature insertion failed:", err);
  } finally {
    // MUST call event.completed(), or Outlook will consider the event “still running”
    event.completed();
  }
}

Office.onReady(() => {
  // Associate manifest function name -> handler
  Office.actions.associate("insertSignature", insertSignature);
});
