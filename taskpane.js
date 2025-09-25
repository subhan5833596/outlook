Office.onReady(function() {
  if (Office.context.mailbox.item) {
    populateFields();
    document.getElementById("Update").onclick = updateUrlInBody;
  }
});

// Async function to populate the fields
async function populateFields() {
  try {
    let item = Office.context.mailbox.item;

    // Get To recipients for Campaign
    const toRecipients = await getRecipientsAsync(item.to);
    document.getElementById("utm_campaign").value = toRecipients.join(",") || "";

    // Source is current user email
    document.getElementById("utm_source").value = Office.context.mailbox.userProfile.emailAddress || "";

    // Medium = email + timestamp
    document.getElementById("utm_medium").value = "email-" + getTimestamp();

    // Content = email subject
    const subject = await getSubjectAsync(item.subject);
    document.getElementById("utm_content").value = subject || "";

    // Term = CC recipients
    const ccRecipients = await getRecipientsAsync(item.cc);
    document.getElementById("utm_term").value = ccRecipients.join(",") || "";
  } catch (err) {
    console.error("Error populating UTM fields:", err);
  }
}

// Helper to get recipients asynchronously
function getRecipientsAsync(getAsyncFunc) {
  return new Promise((resolve, reject) => {
    getAsyncFunc.getAsync(function(res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        const emails = res.value.map(r => r.emailAddress);
        resolve(emails);
      } else {
        reject(res.error);
      }
    });
  });
}

// Helper to get subject asynchronously
function getSubjectAsync(subjectObj) {
  return new Promise((resolve, reject) => {
    subjectObj.getAsync(function(res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(res.error);
    });
  });
}

// Generate timestamp
function getTimestamp() {
  let t = new Date();
  return t.getFullYear() + "-" +
         String(t.getMonth()+1).padStart(2,'0') + "-" +
         String(t.getDate()).padStart(2,'0') + "_" +
         String(t.getHours()).padStart(2,'0') + "-" +
         String(t.getMinutes()).padStart(2,'0');
}

// Update all links in email body with UTM
function updateUrlInBody() {
  const campaign = document.getElementById("utm_campaign").value || "";
  const source = document.getElementById("utm_source").value || "";
  const medium = document.getElementById("utm_medium").value || "";
  const content = document.getElementById("utm_content").value || "";
  const term = document.getElementById("utm_term").value || "";

  const utm = `utm_campaign=${encodeURIComponent(campaign)}&utm_source=${encodeURIComponent(source)}&utm_medium=${encodeURIComponent(medium)}&utm_content=${encodeURIComponent(content)}&utm_term=${encodeURIComponent(term)}`;

  Office.context.mailbox.item.body.getAsync("html", function(res) {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      let body = res.value;

      // Replace all href links with UTM
      body = body.replace(/href="([^"]+)"/g, (match, p1) => {
        try {
          let url = new URL(p1);
          url.search = utm;
          return `href="${url.toString()}"`;
        } catch (e) {
          return match; // skip invalid URLs
        }
      });

      Office.context.mailbox.item.body.setAsync(body, { coercionType: Office.CoercionType.Html }, function(setRes) {
        if (setRes.status === Office.AsyncResultStatus.Succeeded) {
          console.log("✅ All links updated with UTM!");
        } else {
          console.error("❌ Failed updating links:", setRes.error);
        }
      });
    }
  });
}

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
