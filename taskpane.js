Office.onReady(function() {
  if (Office.context.mailbox.item) {
    populateFields();
    document.getElementById("updateUrlBtn").onclick = updateUTMInSignature; // signature only
    document.getElementById("openSigHelpBtn").onclick = () => {
      window.open("https://support.microsoft.com/en-us/office/create-and-add-an-email-signature-in-outlook-for-windows-53f5e0d4-5a23-4a4e-8a90-5d9188e0d1b3", "_blank");
    };
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
function updateUTMInSignature() {
  const campaign = document.getElementById("utm_campaign").value || "";
  const source = document.getElementById("utm_source").value || "";
  const medium = document.getElementById("utm_medium").value || "";
  const content = document.getElementById("utm_content").value || "";
  const term = document.getElementById("utm_term").value || "";

  const utm = `utm_campaign=${encodeURIComponent(campaign)}&utm_source=${encodeURIComponent(source)}&utm_medium=${encodeURIComponent(medium)}&utm_content=${encodeURIComponent(content)}&utm_term=${encodeURIComponent(term)}`;

  Office.context.mailbox.item.body.getAsync("html", function(res) {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      let body = res.value;

      // Find the specific signature link
      const sigUrl = "https://swag.thriftops.com";
      const sigRegex = new RegExp(`<a\\b[^>]*href=["']?${sigUrl}["']?[^>]*>`, "gi");

      if (sigRegex.test(body)) {
        // Replace the link with UTM
        const updatedBody = body.replace(sigRegex, (match) => {
          try {
            let url = new URL(sigUrl);
            url.search = utm;
            return match.replace(sigUrl, url.toString());
          } catch (e) {
            return match; // skip invalid URLs
          }
        });

        // Set updated body
        Office.context.mailbox.item.body.setAsync(updatedBody, { coercionType: Office.CoercionType.Html }, function(setRes) {
          if (setRes.status === Office.AsyncResultStatus.Succeeded) {
            console.log("✅ Signature link updated with UTM!");
          } else {
            console.error("❌ Failed updating signature link:", setRes.error);
          }
        });

      } else {
        console.log("Signature link not found in the body.");
      }

    } else {
      console.error("❌ Could not fetch email body:", res.error);
    }
  });
}





