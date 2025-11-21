console.log('FunctionFile.js loaded');

async function insertSignature(event) {
  console.log('insertSignature called');
  try {
    const sigHtml = buildSignatureHtml();
    console.log('Signature HTML prepared');

    Office.context.mailbox.item.body.prependAsync(
      sigHtml,
      { coercionType: Office.CoercionType.Html },
      (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          console.error('prependAsync failed:', res.error);
        } else {
          console.log('Signature inserted successfully');
        }
        event.completed();
      }
    );
  } catch (e) {
    console.error('insertSignature error:', e);
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
        <a href="https://www.firstenergycorp.com">www.firstenergycorp.com</a><br/>
        <span>Email: user@firstenergycorp.com</span>
      </td></tr>
    </table>
  `;
}
