async function insertSignature(event) {
  try {
    const sigHtml = buildSignatureHtml();

    Office.context.mailbox.item.body.prependAsync(
      sigHtml,
      { coercionType: Office.CoercionType.Html },
      (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("prependAsync failed:", res.error);
        }
        event.completed(); // Always complete so the banner clears
      }
    );
  } catch (e) {
    console.error("insertSignature error:", e);
    event.completed();
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
        https://www.firstenergycorp.comwww.firstenergycorp.com</a><br/>
        <span>Email: user@firstenergycorp.com</span>
      </td></tr>
    </table>
  `;
}