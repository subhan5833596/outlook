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

      // Find signature block inserted by add-in
      const sigRegex = /<!--\s*utm-signature-start\s*-->([\s\S]*?)<!--\s*utm-signature-end\s*-->/i;
      const sigMatch = body.match(sigRegex);

      if (sigMatch) {
        let signatureHTML = sigMatch[1];

        // Replace all href links in signature
        signatureHTML = signatureHTML.replace(/<a\b[^>]*\bhref=["']?([^"'>]+)["']?[^>]*>/gi, (match, url) => {
          try {
            let newUrl = new URL(url, "https://dummybase.com"); // dummy base for relative URLs
            newUrl.search = utm; // replace query string
            return match.replace(url, newUrl.toString().replace("https://dummybase.com", ""));
          } catch (e) {
            return match; // skip invalid URLs
          }
        });

        // Replace the signature block in body
        const newBody = body.replace(sigRegex, `<!-- utm-signature-start -->${signatureHTML}<!-- utm-signature-end -->`);

        // Set updated body
        Office.context.mailbox.item.body.setAsync(newBody, { coercionType: Office.CoercionType.Html }, function(setRes) {
          if (setRes.status === Office.AsyncResultStatus.Succeeded) {
            console.log("✅ Signature links updated with UTM!");
          } else {
            console.error("❌ Failed updating signature links:", setRes.error);
          }
        });

      } else {
        console.log("No signature block found in the body.");
      }

    } else {
      console.error("❌ Could not fetch email body:", res.error);
    }
  });
}

