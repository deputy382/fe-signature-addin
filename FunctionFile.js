// For ExecuteFunction, a global function name must match the manifest's FunctionName.
async function insertSignature(event) {
  try {
    const sigHtml = buildSignatureHtml();

    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (getRes) => {
      if (getRes.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed();
        return;
      }

      const currentHtml = getRes.value || "";
      const marker = "<!-- FE_SIGNATURE_MARKER -->";
      const newHtml = currentHtml.includes(marker)
        ? currentHtml.replace(new RegExp(`${marker}[\\s\\S]*$`, "m"), `${marker}\n${sigHtml}\n`)
        : `${currentHtml}\n${marker}\n${sigHtml}\n`;

      Office.context.mailbox.item.body.setAsync(
        newHtml,
        { coercionType: Office.CoercionType.Html },
        () => event.completed()
      );
    });
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