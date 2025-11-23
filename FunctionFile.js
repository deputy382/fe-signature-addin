
/*
 * FunctionFile.js — FE signature (standard placement) with inline logo (CID)
 * Placement rules:
 *  - New compose: append at bottom
 *  - Reply/Forward: insert just below the reply/forward header
 *
 * Notes:
 *  - Inline image via CID (cid:fe-logo.png).
 *  - Marker prevents duplicate inserts.
 *  - Consolidated Office.onReady(...) with association + diagnostics.
 */

console.log("FunctionFile.js loaded");

const MARKER = "FE_SIGNATURE_MARKER";
const INLINE_IMAGE_NAME = "fe-logo.png"; // CID name used in cid:...

// TODO: Replace with your actual Base64 PNG for the FE logo (transparent, ~40–42px height recommended).
const FE_LOGO_BASE64 = "<BASE64_PNG_STRING_FOR_FE_LOGO>";

/* ========================= Helpers ========================= */

function getBodyHtmlAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      (res) =>
        res.status === Office.AsyncResultStatus.Succeeded
          ? resolve(res.value || "")
          : reject(res.error)
    );
  });
}

function setBodyHtmlAsync(newHtml) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(
      newHtml,
      { coercionType: Office.CoercionType.Html },
      (res) =>
        res.status === Office.AsyncResultStatus.Succeeded
          ? resolve()
          : reject(res.error)
    );
  });
}

/**
 * Insert signature just after common reply/forward header markers.
 * If none are found, append signature at bottom.
 */
function insertAfterReplyHeader(html, sigHtml) {
  const patterns = [
    /<div[^>]*id=["']divRplyFwdMsg["'][^>]*>/i,         // Classic Outlook wrapper
    /<hr[^>]*>/i,                                       // Horizontal rule before quoted content
    /<div[^>]*>.*?On .*? wrote:\s*<\/div>/is,           // English textual header in a div
    /On .*? wrote:/i,                                   // English textual header (fallback)
    /<blockquote[^>]*>/i,                               // Quoted block (OWA/New Outlook)
    /<div[^>]*class=["'][^"']*(gmail_quote|moz-cite-prefix|yahoo_quoted|WordSection1)["'][^>]*>/i
  ];
  for (const pattern of patterns) {
    const match = html.match(pattern);
    if (match) {
      const idx = html.indexOf(match[0]) + match[0].length;
      return html.slice(0, idx) + "\n" + sigHtml + "\n" + html.slice(idx);
    }
  }
  // Fallback: append at bottom
  return `${html}\n${sigHtml}`;
}

/**
 * Attach inline image from Base64 and return the cid:... tag.
 */
function attachInlineImageAndGetCidTag(base64Png) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      base64Png,
      INLINE_IMAGE_NAME,
      { isInline: true },
      (attachResult) => {
        if (attachResult.status !== Office.AsyncResultStatus.Succeeded) {
          return reject(attachResult.error);
        }
        // Render with fixed max-height to align nicely with text block.
        const imgTag = `<id:${INLINE_IMAGE_NAME}`;
        resolve(imgTag);
      }
    );
  });
}

/* ========================= Signature builder ========================= */
/*
   Layout (matches your screenshot):
   ┌───────────────┬──────────────────────────────────────────────────────┐
   │     Logo      │ Name (bold)                                          │
   │               │ Title                                                │
   │               │ office: … | cell: …                                  │
   │               │ email (blue link)                                    │
   │               │ address | mailstop … / site                          │
   └───────────────┴──────────────────────────────────────────────────────┘
*/
function buildSignatureHtml(imgTag, person) {
  const nameLine   = `<div style="font-size:13px; font-weight:600; color:#000;">${person.displayName}</div>`;
  const titleLine  = `<div style="color:#000;">${person.title}</div>`;
  const phoneLine  = `<div style="color:#000;">office: ${person.officePhone}${person.officeExt ? ` (${person.officeExt})` : ""} | cell: ${person.mobile}</div>`;
  const emailLine  = `<div>mailto:${person.email}${person.email}</a></div>`;
  const addrLine   = `<div style="color:#000;">${person.address} | mailstop: ${person.mailstop}${person.site ? ` / ${person.site}` : ""}</div>`;

  return `
    <!-- ${MARKER} -->
    <table role="presentation" style="font-family:'Segoe UI', Arial, sans-serif; font-size:12px; line-height:1.35;">
      <tr>
        <td style="vertical-align:top; padding-right:12px;">
          ${imgTag}
        </td>
        <td style="vertical-align:top; padding:2px 0;">
          ${nameLine}
          ${titleLine}
          ${phoneLine}
          ${emailLine}
          ${addrLine}
        </td>
      </tr>
    </table>
  `.trim();
}

/* ========================= Event-based handler ========================= */
async function insertSignature(event) {
  try {
    const t0 = performance.now();

    const bodyHtml = await getBodyHtmlAsync();

    // Idempotence: skip if already present
    const hasMarker =
      bodyHtml.includes(`<!-- ${MARKER} -->`) ||
      bodyHtml.includes(MARKER);
    if (hasMarker) {
      console.log("Signature already present; skipping.");
      return;
    }

    // v1: static user block (we’ll replace with Graph data later)
    const person = {
      displayName: "Shane Francis",
      title: "Systems Administrator B IV",
      officePhone: "3303238382",
      officeExt: "850-2601",
      mobile: "330-323-8382",
      email: "sfrancis@firstenergycorp.com",
      address: "341 White Pond Drive, Akron, OH 44320",
      mailstop: "A-FEHQ-A2",
      site: "Akron FirstEnergy Headquarters"
    };

    // Attach logo inline and build the signature HTML
    const imgCidTag = await attachInlineImageAndGetCidTag(FE_LOGO_BASE64);
    const sigHtml   = buildSignatureHtml(imgCidTag, person);

    // Standard placement (below reply header / else bottom)
    const newHtml = insertAfterReplyHeader(bodyHtml, sigHtml);
    await setBodyHtmlAsync(newHtml);

    const t1 = performance.now();
    console.log(`Signature inserted at standard location. (${Math.round(t1 - t0)} ms)`);
  } catch (err) {
    console.error("Signature insertion failed:", err);
  } finally {
    // Must signal completion for event-based activation
    event.completed();
  }
}

/* ========================= Consolidated Office.onReady ========================= */
Office.onReady(async () => {
  // Verify the autorun/runtime loaded
  console.log("Autorun runtime loaded:", Office.context.platform);

  // REQUIRED: wire manifest function -> handler
  Office.actions.associate("insertSignature", insertSignature);

  // Optional diagnostics: what compose type are we in? (NewMail | Reply | Forward)
  if (Office.context?.mailbox?.item?.getComposeTypeAsync) {
    Office.context.mailbox.item.getComposeTypeAsync((res) => {
      console.log("ComposeType:", res.status, res.value);
    });
  }

  // Keep this block lightweight—heavy init can slow event handlers.
});
