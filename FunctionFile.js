Office.onReady(() => {
  Office.actions.associate("insertSignature", insertSignature);
});

async function insertSignature(event) {
  try {
    const currentHtml = await getBodyHtmlAsync();
    const signatureHtml = buildSignatureHtml();
    const marker = "<!-- FE_SIGNATURE_MARKER -->";

    let newHtml;
    if (currentHtml.includes(marker)) {
      newHtml = currentHtml.replace(
        new RegExp(`${marker}[\\s\\S]*$`, "m"),
        `${marker}\n${signatureHtml}\n`
      );
    } else {
      newHtml = `${currentHtml}\n${marker}\n${signatureHtml}\n`;
    }

    await setBodyHtmlAsync(newHtml);
    event.completed();
  } catch (e) {
    console.error("insertSignature error:", e);
    event.completed({ error: e.message || "Failed to insert signature." });
  }
}

function buildSignatureHtml() {
  return `
    <!-- FE_SIGNATURE_MARKER -->
    <table style="font-family:Segoe UI, Arial; font-size:12px; line-height:1.35;">
      <tr><td style="padding:6px 0;">
        <strong>FirstEnergy</strong><br/>
        Employee Name | Title<br/>
        Department<br/>
        <a href="https://www.firstenergycorpirstenergycorp.com</a><br/>
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