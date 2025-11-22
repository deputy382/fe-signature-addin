
console.log("FunctionFile.js loaded");

Office.onReady(() => {
  // Associate manifest function "insertSignature" with this handler.
  Office.actions.associate("insertSignature", insertSignature);
});

/**
 * Event-based handler: inserts FE signature if not already present.
 * Must call event.completed() when finished.
 */
async function insertSignature(event) {
  try {
    // Read current body as HTML
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value || "");
          } else {
            reject(result.error);
          }
        }
      );
    }).then(async (bodyHtml) => {
      // Use a robust marker check (HTML comment or a unique token in signature)
      const MARKER = "FE_SIGNATURE_MARKER";
      const alreadyHasSignature =
        bodyHtml.includes("<!-- " + MARKER + " -->") ||
        bodyHtml.includes(MARKER);

      if (alreadyHasSignature) {
        console.log("FE signature marker found; skipping insert.");
        return;
      }

      // Build REAL HTML (no entity escaping)
      const sigHtml = `
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

      // Prepend our signature at the top of the compose body.
      await new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.prependAsync(
          sigHtml,
          { coercionType: Office.CoercionType.Html },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              resolve();
            } else {
              reject(asyncResult.error);
            }
          }
        );
      });

      console.log("Signature inserted.");
    });
  } catch (err) {
    console.error("Signature insertion failed:", err);
  } finally {
    // MUST call event.completed() regardless of success/failure
    event.completed();
  }
}
