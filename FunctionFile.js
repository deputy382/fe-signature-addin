console.log('FunctionFile.js loaded');


Office.onReady(() => {
  window.insertSignature = async function(event) {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const bodyHtml = result.value;
        if (!bodyHtml.includes('FE_SIGNATURE_MARKER')) {
          // Insert signature
          const sigHtml = `
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
          Office.context.mailbox.item.body.prependAsync(
            sigHtml,
            { coercionType: Office.CoercionType.Html },
            () => event.completed()
          );
        } else {
          // Signature already present, do nothing
          event.completed();
        }
      } else {
        // Could not read body, do nothing
        event.completed();
      }
    });
  };
});
