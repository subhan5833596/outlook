Office.onReady(() => {
  if (!Office.context.mailbox.item) return;

  let item = Office.context.mailbox.item;

  // AUTO-FILL FIELDS
  document.getElementById("utm_campaign").value = (item.to && item.to.length) ? item.to.map(t=>t.emailAddress).join(", ") : "";
  document.getElementById("utm_source").value = item.from ? item.from.emailAddress : "";
  document.getElementById("utm_medium").value = "email-" + new Date().toISOString().replace(/[:T]/g, "-").slice(0,16);
  document.getElementById("utm_content").value = item.subject || "";
  document.getElementById("utm_term").value = "";

  // UPDATE URL BUTTON
  document.getElementById("updateUrlBtn").onclick = updateUrlInSignature;

  // OPEN SIGNATURE HELP
  document.getElementById("openSigHelpBtn").onclick = () => {
    window.open("https://support.microsoft.com/en-us/office/create-and-add-an-email-signature-in-outlook-4e79d2eb-0f5f-4d60-bf29-0e9a4f3133b9", "_blank");
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
