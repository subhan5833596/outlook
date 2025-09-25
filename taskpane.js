Office.onReady(() => {
  if (!Office.context.mailbox.item) return;
  let item = Office.context.mailbox.item;

  // Auto-fill fields
  const campaignEl = document.getElementById("utm_campaign");
  const sourceEl = document.getElementById("utm_source");
  const mediumEl = document.getElementById("utm_medium");
  const contentEl = document.getElementById("utm_content");
  const termEl = document.getElementById("utm_term");

  campaignEl.value = (item.to && item.to.length) ? item.to.map(t => t.emailAddress).join(", ") : "";
  sourceEl.value = (item.from && item.from.emailAddress) ? item.from.emailAddress : "";
  mediumEl.value = "email-" + new Date().toISOString().replace(/[:T]/g, "-").slice(0,16);
  contentEl.value = item.subject || "";
  termEl.value = "";

  // Update URL button
  document.getElementById("updateUrlBtn").onclick = updateUrlInSignature;

  // ✅ Attach the signature help button here
  document.getElementById("openSigHelpBtn").onclick = () => {
    window.open("https://support.microsoft.com/en-us/office/create-and-add-an-email-signature-in-outlook-for-windows-53f5e0d4-5a23-4a4e-8a90-5d9188e0d1b3", "_blank");
  };
});


// UPDATE ALL LINKS IN BODY WITH UTM
function updateUrlInSignature() {
  const campaign = document.getElementById("utm_campaign").value;
  const source = document.getElementById("utm_source").value;
  const medium = document.getElementById("utm_medium").value;
  const content = document.getElementById("utm_content").value;
  const term = document.getElementById("utm_term").value;

  const utm = { utm_campaign: campaign, utm_source: source, utm_medium: medium, utm_content: content };
  if(term) utm.utm_term = term;

  Office.context.mailbox.item.body.getAsync("html", res => {
    if(res.status !== Office.AsyncResultStatus.Succeeded) return alert("❌ Could not read body: " + res.error.message);

    let body = res.value;

    // Replace all links
    body = body.replace(/href="([^"]+)"/g, (match, urlStr) => {
      try {
        let url = new URL(urlStr);
        Object.keys(utm).forEach(k => url.searchParams.set(k, utm[k]));
        return `href="${url.toString()}"`;
      } catch(e) {
        return match; // skip invalid URLs
      }
    });

    Office.context.mailbox.item.body.setAsync(body, { coercionType: "html" }, setRes => {
      if(setRes.status === Office.AsyncResultStatus.Succeeded) {
        console.log("✅ All signature links updated with UTM!");
        Office.context.mailbox.item.notificationMessages.replaceAsync("utmMsg", {
          type: "informationalMessage",
          message: "✅ All links updated with UTM!",
          icon: "Icon.16x16",
          persistent: false
        });
      } else {
        console.error("❌ Failed updating links:", setRes.error);
        alert("❌ Failed updating links: " + setRes.error.message);
      }
    });
  });
}
