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
      console.log(body);
      // Regex to match all <a href="..."> links
      body = body.replace(/<a\b[^>]*\bhref=["']?([^"'>]+)["']?[^>]*>/gi, (match, url) => {
        try {
          let newUrl = new URL(url, "https://dummybase.com"); // dummy base for relative URLs
          newUrl.search = utm; // append UTM
          return match.replace(url, newUrl.toString().replace("https://dummybase.com", ""));
        } catch (e) {
          return match; // skip invalid URLs
        }
      });

      Office.context.mailbox.item.body.setAsync(body, { coercionType: Office.CoercionType.Html }, function(setRes) {
        if (setRes.status === Office.AsyncResultStatus.Succeeded) {
          console.log("✅ All links in signature updated with UTM!");
        } else {
          console.error("❌ Failed updating links:", setRes.error);
        }
      });

    } else {
      console.error("❌ Could not fetch email body:", res.error);
    }
  });
}
let quill;

// Open modal + init Quill
document.getElementById("openEditorBtn").onclick = () => {
  document.getElementById("editorModal").style.display = "block";

  if (!quill) {
    quill = new Quill('#quillEditor', {
      theme: 'snow',
      modules: {
        toolbar: [
          ['bold', 'italic', 'underline', 'strike'],
          [{ 'font': [] }, { 'size': [] }],
          [{ 'color': [] }, { 'background': [] }],
          [{ 'align': [] }],
          ['link', 'image'],
          [{ 'list': 'ordered' }, { 'list': 'bullet' }]
        ]
      }
    });

    // Load saved signature into editor
    const savedSig = localStorage.getItem("customSignature") || "";
    if (savedSig) {
      quill.root.innerHTML = savedSig;
    }
  }
};

// Close modal
function closeEditor() {
  document.getElementById("editorModal").style.display = "none";
}

document.getElementById("saveSignatureBtn").onclick = () => {
  const sigHTML = quill.root.innerHTML;
  localStorage.setItem("customSignature", sigHTML);

  // Show success message instead of alert
  document.getElementById("statusMsg").innerText = "✅ Signature saved!";
  
  // Auto-hide modal after 1 sec
  setTimeout(() => {
    document.getElementById("editorModal").style.display = "none";
    document.getElementById("statusMsg").innerText = "";
  }, 1000);
};


function insertSignature() {
  const sigHTML = localStorage.getItem("customSignature") || "";
  if (!sigHTML) {
    document.getElementById("statusMsg").innerText = "⚠️ No signature saved yet!";
    return;
  }

  Office.context.mailbox.item.body.setSelectedDataAsync(
    sigHTML,
    { coercionType: Office.CoercionType.Html },
    (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        document.getElementById("statusMsg").innerText = "✅ Signature inserted into email!";
      } else {
        document.getElementById("statusMsg").innerText = "❌ Failed: " + res.error.message;
        console.error(res.error);
      }
    }
  );
}







