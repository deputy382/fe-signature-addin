// Optional: basic init hook (not required for event-based, but harmless)
Office.initialize = function () {};

// Simple namespace for your signature logic
window.FESignature = (() => {
  // Build a basic signature; swap this with your branded HTML as needed.
  function buildSignatureHtml() {
    return `
      <!-- FE_SIGNATURE_MARKER -->
      <table style="font-family:Segoe UI, Arial; font-size:12px; line-height:1.35;">
        <tr><td style="padding:6px 0;">
          <strong>FirstEnergy</strong><br/>
          Employee Name | Title<br/>
          Department<br/>
          https://www.firstenergycorp.comfirstenergycorp.com</a><br/>
          <span>Email: user@firstenergycorp.com</span>
        </td></tr>
      </table>
    `;
  }

  function setBodyHtmlAsync(html) {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.setAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        (res) => res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error)
      );
    });
  }

  function getBodyHtmlAsync() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        (res) => res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value || "") : reject(res.error)
      );
    });
  }

  // Inserts (or refreshes) the signature at the end of the compose body.
  async function insertSignatureOnCompose() {
    const currentHtml = await getBodyHtmlAsync();
    const signatureHtml = buildSignatureHtml();

    const marker = "<!-- FE_SIGNATURE_MARKER -->";
    let newHtml;

    if (currentHtml.includes(marker)) {
      // Replace anything after the marker with a fresh signature block
      newHtml = currentHtml.replace(
        new RegExp(`${marker}[\\|\\n|\\r|\\s|\\S]*$`, "m"),
        `${marker}\n<div>${signatureHtml}</div>\n`
      );
    } else {
      // Append a marker + signature at the end
      newHtml = `${currentHtml}\n${marker}\n<div>${signatureHtml}</div>\n`;
    }

    await setBodyHtmlAsync(newHtml);
  }

  // Optional: notification helper
  function statusUpdate(icon, text) {
    try {
      Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        icon: icon || "icon16",
        message: text || "Operation complete.",
        persistent: false
      });
    } catch (_) { /* no-op in event runtime */ }
  }

  return { insertSignatureOnCompose, statusUpdate };
