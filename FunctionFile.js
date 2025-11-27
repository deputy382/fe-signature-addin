/* FunctionFile.js — FE Signature (button-driven, Entra/Graph SSO, no static fallback) */

/* ========= CONFIG: EDIT THESE ========= */
const API_BASE = "https://api.your-backend.com/fe-signature // TODO: Update to your backend endpoint"; // TODO: change to your backend host
const COMPANY_TEXT = "FirstEnergy Corp."; // optional default if backend doesn't supply
const logoUrl = "https://yourcompany.com/logo.png // TODO: Replace with actual logo URL"; // TODO: swap to actual logo URL

/* ========= INTERNAL ========= */
const SIG_MARKER = "FE_SIGNATURE_MARKER";
const SIG_COMMENT = `<!-- ${SIG_MARKER} -->`;

/**
 * Build the signature HTML from user data.
 * @param {object} d Normalized DTO fields expected from backend:
 *  {
 *    displayName, jobTitle, mail,
 *    mobilePhone, businessPhone,
 *    officeLocation, mailStop,
 *    company? // optional; will default to COMPANY_TEXT if absent
 *  }
 */
function buildSignatureHtml(d) {
  const esc = (s) => (s || "").toString()
    .replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;",">" :"&gt;" ,'"' :"&quot;" }[c])); const displayName=esc(d.displayName); const jobTitle=esc(d.jobTitle); const mail=esc(d.mail); const mobilePhone=esc(d.mobilePhone); const businessPhone=esc(d.businessPhone); const office=esc(d.officeLocation); const mailStop=esc(d.mailStop); const company=esc(d.company || COMPANY_TEXT); return ( SIG_COMMENT + `
	<table cellpadding="0" cellspacing="0" style="font-family:'Segoe UI', Arial, sans-serif; font-size:12px; line-height:1.35;">
		<tr>
			<td style="vertical-align:middle; padding-right:16px;">
				<img src="${logoUrl}" alt="FirstEnergy Logo" style="border-left:2px solid #003366; padding-left:16px; verticaldiv>
      <div style=" font-size:13px; color:#222; margin-bottom:6px;">${jobTitle}< div>
				<div style="margin-bottom:2px;">
					<span style="color:#0072c6;">office: ${businessPhone}</span>
					<span style="color:#222;"> | </span>
					<span style="color:#0072c6;">cell: ${mobilePhone}</span>
				</div>
				<div style="margin-bottom:2px;">
					<a href="mailto{mail}</a>
      </div>
      <div style=" color:#0072c6;"> ${office} | mailstop: ${mailStop} ${company}
				</div>
			</td>
		</tr>
	</table>
    `.trim()
  );
}

/**
 * Get user data from your backend via SSO → OBO → Graph.
 * Backend endpoint: GET /api/signature/me
 * - Validates Office SSO token
 * - Exchanges via OBO for Graph token
 * - Calls Graph /me (and optionally extensions)
 * - Returns normalized DTO
 */
async function fetchUserData() {
  // 1) Get Office SSO token (scoped to your manifest's Resource/Application ID URI)
  const ssoToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });

  // 2) Call backend with SSO token
  const res = await fetch(`${API_BASE}/api/signature/me`, {
    method: "GET",
    headers: { Authorization: `Bearer ${ssoToken}` }
  });
  if (!res.ok) throw new Error(`API ${res.status}`);

  const dto = await res.json();

  // 3) Normalize defensively (some props may be missing)
  const businessPhone = Array.isArray(dto.businessPhones) && dto.businessPhones.length
    ? dto.businessPhones[0]
    : (dto.businessPhone || "");

  return {
    displayName:   dto.displayName || "",
    jobTitle:      dto.jobTitle || "",
    mail:          dto.mail || dto.userPrincipalName || "",
    mobilePhone:   dto.mobilePhone || "",
    businessPhone,
    officeLocation:dto.officeLocation || "",
    // If you store mail stop in onPremisesExtensionAttributes.extensionAttributeX, map it server-side,
    // or keep this fallback client-side until you finalize which attribute you use:
    mailStop:      dto.mailStop || (dto.onPremisesExtensionAttributes?.extensionAttribute1 || ""),
    company:       dto.company || COMPANY_TEXT
  };
}

/** Insert signature using Outlook's signature API (host-managed placement) */
async function doInsertSignature() {
  const data = await fetchUserData();
  const sigHtml = buildSignatureHtml(data);

  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSignatureAsync(sigHtml, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        console.error("setSignatureAsync failed:", res.error);
        reject(res.error);
      }
    });
  });
}

/** Ribbon command entry point (manifest ExecuteFunction → insertSignature) */
async function insertSignature(event) {
  try { await doInsertSignature(); }
  catch (err) { console.error("insertSignature failed:", err); }
  finally { if (event && typeof event.completed === "function") event.completed(); }
}

// Expose for manifest
window.insertSignature = insertSignature;

// Optional logging
Office.onReady(() => {
  console.log("FE Signature FunctionFile ready:", Office.context.platform);
});
