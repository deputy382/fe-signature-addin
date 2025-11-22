console.log('FunctionFile.js loaded');

// Expose the function globally for ExecuteFunction in the manifest
window.insertSignature = async function(event) {
  console.log('insertSignature called');
  try {
    const sigHtml = buildSignatureHtml();
    console.log('Signature HTML prepared');

    Office.context.mailbox.item.body.prependAsync(
      sigHtml,
      { coercionType: Office.CoercionType.Html },
      (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          console.error('prependAsync failed:', res.error?.message || res.error || res);
        } else {
          console.log('Signature inserted successfully');
        }
        // Always complete so the banner clears
        event.completed();
      }
    );
  } catch (e) {
    console.error('insertSignature error:', e?.message || e);
    event.completed();
  }
};

function buildSignatureHtml() {
  // IMPORTANT: real HTML tags here (no &lt; / &gt; entities)
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