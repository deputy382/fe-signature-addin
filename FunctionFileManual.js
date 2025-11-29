function insertSignature(event) {
    var signatureHtml = "<p style='font-family:Arial; font-size:12pt;'>Regards,<br><strong>FirstEnergy</strong><br>www.firstenergycorp.com</p>";
    Office.context.mailbox.item.body.setSelectedDataAsync(signatureHtml, { coercionType: "html" }, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Signature inserted successfully.");
        } else {
            console.error("Failed to insert signature: " + result.error.message);
        }
        if (event && typeof event.completed === "function") event.completed();
    });
}
window.insertSignature = insertSignature;
