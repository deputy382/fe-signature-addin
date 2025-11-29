/* FunctionFileManual.js — FE Signature (commands + autorun) */
console.log('FunctionFile.js loaded');

// ... [all your helper functions and signature logic] ...

// Ribbon button
async function insertSignature(event) {
  try { await doInsertSignature(); }
  catch (err) { console.error('❌ insertSignature failed:', err); }
  finally { if (event && typeof event.completed === 'function') event.completed(); }
}

// Expose for ExecuteFunction (Commands v1.0)
window.insertSignature = insertSignature;

// Office.onReady for autorun, etc.
Office.onReady(() => {
  console.log('Autorun runtime loaded:', Office.context.platform);
  try { Office.actions.associate('onNewCompose', onNewCompose); } catch (e) { /* already bound */ }
});
