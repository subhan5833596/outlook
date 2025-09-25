// Save signature into email body
function saveSignature() {
  const sig = document.getElementById("signatureContent").value;

  Office.context.mailbox.item.body.setAsync(
    sig,
    { coercionType: Office.CoercionType.Html, asyncContext: "set" },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        alert("✅ Signature added to email!");
      } else {
        alert("❌ Error: " + asyncResult.error.message);
      }
    }
  );
}

// Update existing links in signature with UTM
function updateUTM() {
  const campaign = document.getElementById("utm_campaign").value;
  const source = document.getElementById("utm_source").value;
  const medium = document.getElementById("utm_medium").value;
  const content = document.getElementById("utm_content").value;
  const term = document.getElementById("utm_term").value;

  // Build UTM query string
  let utmParams = `utm_campaign=${campaign}&utm_source=${source}&utm_medium=${medium}&utm_content=${content}&utm_term=${term}`;

  Office.context.mailbox.item.body.getAsync(
    "html",
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        let body = asyncResult.value;

        // Replace any existing URL with UTM
        body = body.replace(/href="([^"]+)"/g, (match, p1) => {
          let newUrl = p1.split("?")[0] + "?" + utmParams;
          return `href="${newUrl}"`;
        });

        Office.context.mailbox.item.body.setAsync(
          body,
          { coercionType: "html" },
          function (res) {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              alert("✅ Signature URLs updated with UTM!");
            } else {
              alert("❌ Update failed: " + res.error.message);
            }
          }
        );
      } else {
        alert("❌ Could not read body: " + asyncResult.error.message);
      }
    }
  );
}
