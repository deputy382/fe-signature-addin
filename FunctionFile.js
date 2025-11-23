
/*
 * FunctionFile.js — FE signature (standard placement) without logo
 * Placement rules:
 *  - New compose: append at bottom
 *  - Reply/Forward: insert just below the reply/forward header
 *
 * Notes:
 *  - No inline image work (logo deferred).
 *  - Hidden marker prevents duplicate inserts.
 *  - A single Office.onReady associates the event handler and logs diagnostics.
 */

console.log("FunctionFile.js loaded");

const MARKER = "FE_SIGNATURE_MARKER";

/* ========================= Helpers ========================= */

function getBodyHtmlAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          resolve(res.value || "");
        } else {
          console.error("getAsync(body HTML) failed:", res.error);
          reject(res.error);
        }
      }
    );
  });
}

function setBodyHtmlAsync(newHtml) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(
      newHtml,
      { coercionType: Office.CoercionType.Html },
      (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          console.error("setAsync(body HTML) failed:", res.error);
          reject(res.error);
        }
      }
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

/* ========================= Signature builder ========================= */
/*
   Layout (matches your screenshot, without logo):
   ┌───────────────────────────────────────────────────────────────┐
   │ Name (bold)                                                   │
   │ Title                                                         │
   │ office: … | cell: …                                           │
   │ email (blue link)                                             │
   │ address | mailstop … / site                                   │
   └───────────────────────────────────────────────────────────────┘
*/
function buildSignatureHtml(person) {
  const nameLine   = `<div style="font-size:13px; font-weight:600; color:#000;">${person.displayName}</div>`;
  const titleLine  = `<div style="color:#000;">${person.title}</div>`;
  const phoneLine  = `<div style="color:#000;">office: ${person.officePhone}${person.officeExt ? ` (${person.officeExt})` : ""} | cell: ${person.mobile}</div>`;
  const emailLine  = `<div>mailto:${person.email}${person.email}</a></div>`;
  const addrLine   = `<div style="color:#000;">${person.address} | mailstop: ${person.mailstop}${person.site ? ` / ${person.site}` : ""}</div>`;

  return `
    <!-- ${MARKER} -->
    <div style="font-family:'Segoe UI', Arial, sans-serif; font-size:12px; line-height:1.35;">
      ${nameLine}
      ${titleLine}
      ${phoneLine}
      ${emailLine}
      ${addrLine}
    </div>
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

    // Build signature and place in standard location
    const sigHtml = buildSignatureHtml(person);
    const newHtml = insertAfterReplyHeader(bodyHtml, sigHtml);

    // Write into body
    await setBodyHtmlAsync(newHtml);

    const t1 = performance.now();
    console.log(`Signature inserted at standard location. (${Math.round(t1 - t0)} ms)`);
  } catch (err) {
    console.error("❌ Signature insertion failed:", err);
  } finally {
    // Must signal completion for event-based activation
    event.completed();
  }
}

/* ========================= Consolidated Office.onReady ========================= */
Office.onReady(async () => {
  console.log("Autorun runtime loaded:", Office.context.platform);

  // REQUIRED: bind manifest function -> handler
  Office.actions.associate("insertSignature", insertSignature);

  // Optional: confirm compose type (NewMail | Reply | Forward)
  if (Office.context?.mailbox?.item?.getComposeTypeAsync) {
    Office.context.mailbox.item.getComposeTypeAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        console.log("ComposeType:", res.value);
      } else {
        console.warn("ComposeType failed:", res.error);
      }
    });
  }
