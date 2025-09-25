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

Office.initialize = function () {
  Office.onReady(function () {
    if (Office.context.mailbox.item) {
      // Autofill values
      let item = Office.context.mailbox.item;

      document.getElementById("campaign").value = item.normalizedSubject || "";
      document.getElementById("source").value = item.from ? item.from.emailAddress : "";
      document.getElementById("medium").value = "email-" + new Date().toISOString().replace(/[:T]/g, "-").slice(0,16);
      document.getElementById("content").value = item.subject || "";
      document.getElementById("term").value = "";
    }
  });
};
function updateUrlInSignature() {
  let campaign = document.getElementById("campaign").value;
  let source = document.getElementById("source").value;
  let medium = document.getElementById("medium").value;
  let content = document.getElementById("content").value;
  let term = document.getElementById("term").value;

  // Build UTM string
  let utm = `utm_campaign=${encodeURIComponent(campaign)}&utm_source=${encodeURIComponent(source)}&utm_medium=${encodeURIComponent(medium)}&utm_content=${encodeURIComponent(content)}`;
  if (term) utm += `&utm_term=${encodeURIComponent(term)}`;

  // Get current body (with signature)
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      let body = res.value;

      // Find first link in signature (e.g. swag.thriftops.com)
      let updatedBody = body.replace(/href="([^"]+)"/, (match, p1) => {
        let url = new URL(p1);
        url.search = utm; // replace query string with new UTM
        return `href="${url.toString()}"`;
      });

      // Update back to mail body
      Office.context.mailbox.item.body.setAsync(updatedBody, { coercionType: Office.CoercionType.Html }, (setRes) => {
        if (setRes.status === Office.AsyncResultStatus.Succeeded) {
          console.log("✅ Signature link updated with UTM params!");
        } else {
          console.error("❌ Failed updating signature:", setRes.error);
        }
      });
    }
  });
}

