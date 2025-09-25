// extra.js
// Loads after taskpane.js. Minimal, safe.
Office.onReady(function() {
  // Attach handlers (Update handler is already bound in taskpane.js to #Update)
  var editBtn = document.getElementById("EditSignatureBtn");
  var modal = document.getElementById("signatureModal");
  var saveBtn = document.getElementById("SaveSig");
  var cancelBtn = document.getElementById("CancelSig");
  var editor = document.getElementById("signatureEditor");

  if (editBtn) editBtn.addEventListener("click", openEditor);
  if (cancelBtn) cancelBtn.addEventListener("click", closeEditor);
  if (saveBtn) saveBtn.addEventListener("click", saveSignature);

  // Try to prefill editor with last-saved signature from localStorage
  function openEditor() {
    var last = localStorage.getItem("utm_signature_html");
    editor.value = last || "";
    modal.style.display = "flex";
  }

  function closeEditor() {
    modal.style.display = "none";
  }

  function saveSignature() {
    var html = editor.value || "";
    // Save locally for next time
    try { localStorage.setItem("utm_signature_html", html); } catch(e){ console.warn("localStorage failed", e); }

    // Insert/update signature into current message body.
    // We'll remove any existing signature that was inserted previously by this add-in,
    // which we mark with <!-- utm-signature-start --> ... <!-- utm-signature-end -->
    Office.context.mailbox.item.body.getAsync("html", function(getRes) {
      if (getRes.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Could not get body:", getRes.error);
        alert("Error reading message body: " + (getRes.error && getRes.error.message));
        return closeEditor();
      }

      var body = getRes.value || "";
      // Remove existing marked signature block if present
      var cleaned = body.replace(/<!--\s*utm-signature-start\s*-->[\s\S]*?<!--\s*utm-signature-end\s*-->/gi, "");

      // Append signature at end inside markers
      var signatureBlock = "<!-- utm-signature-start -->" + html + "<!-- utm-signature-end -->";

      // If body contains </body>, insert before it; otherwise append to end
      if (/<\/body>/i.test(cleaned)) {
        cleaned = cleaned.replace(/<\/body>/i, signatureBlock + "</body>");
      } else {
        cleaned = cleaned + signatureBlock;
      }

      Office.context.mailbox.item.body.setAsync(cleaned, { coercionType: Office.CoercionType.Html }, function(setRes) {
        if (setRes.status === Office.AsyncResultStatus.Succeeded) {
          // show small notification
          try {
            Office.context.mailbox.item.notificationMessages.replaceAsync("utmSigMsg", {
              type: "informationalMessage",
              message: "Signature updated",
              icon: "Icon.16x16",
              persistent: false
            });
          } catch (e) {
            console.log("notify skipped", e);
          }
        } else {
          console.error("setAsync failed:", setRes.error);
          alert("Failed to insert signature: " + (setRes.error && setRes.error.message));
        }
        closeEditor();
      });
    });
  }
});
