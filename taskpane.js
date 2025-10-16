// Office.onReady(function () {
//   if (Office.context.mailbox.item) {
//     populateFields();
//     document.getElementById("updateUrlBtn").onclick = updateUTMInSignature; // signature only
//   }
// });

// ✅ Step 6: Load UTM Manager only for supported accounts and mail compose items
Office.onReady(function (info) {
  if (!Office.context.mailbox || !Office.context.mailbox.item) {
    console.warn("⚠️ Office context not ready yet — skipping init");
    return;
  }

  if (info.host === Office.HostType.Outlook) {
    const mailbox = Office.context.mailbox;

    // Skip unsupported item types (e.g., meeting invites)
    const itemType = mailbox.item?.itemType;
    if (itemType && itemType !== Office.MailboxEnums.ItemType.Message) {
      console.log("🚫 Not a mail compose item — UTM Manager disabled.");
      return;
    }

    // Skip unsupported account types (adjust regex as per your environment)
    const userEmail = mailbox.userProfile?.emailAddress || "";
    if (!/@(outlook\.com|office365\.com|yourcompany\.com)$/i.test(userEmail)) {
      console.log("🚫 Unsupported account detected:", userEmail);
      return;
    }

    // ✅ All checks passed — initialize UTM Manager
    console.log(
      "✅ Supported account and item type — initializing UTM Manager."
    );
    populateFields();
    document.getElementById("updateUrlBtn").onclick = updateUTMInSignature;
  }
});

// Async function to populate the fields
async function populateFields() {
  try {
    let item = Office.context.mailbox.item;

    // Get To recipients for Campaign
    const toRecipients = await getRecipientsAsync(item.to);
    document.getElementById("utm_campaign").value =
      toRecipients.join(",") || "";

    // Source is current user email
    document.getElementById("utm_source").value =
      Office.context.mailbox.userProfile.emailAddress || "";

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
    getAsyncFunc.getAsync(function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        const emails = res.value.map((r) => r.emailAddress);
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
    subjectObj.getAsync(function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(res.error);
    });
  });
}

// Generate timestamp
function getTimestamp() {
  let t = new Date();
  return (
    t.getFullYear() +
    "-" +
    String(t.getMonth() + 1).padStart(2, "0") +
    "-" +
    String(t.getDate()).padStart(2, "0") +
    "_" +
    String(t.getHours()).padStart(2, "0") +
    "-" +
    String(t.getMinutes()).padStart(2, "0")
  );
}

// function updateUTMInSignature() {
//   const campaign = document.getElementById("utm_campaign").value || "";
//   const source = document.getElementById("utm_source").value || "";
//   const medium = document.getElementById("utm_medium").value || "";
//   const content = document.getElementById("utm_content").value || "";
//   const term = document.getElementById("utm_term").value || "";

//   const utm = `utm_campaign=${encodeURIComponent(campaign)}&utm_source=${encodeURIComponent(source)}&utm_medium=${encodeURIComponent(medium)}&utm_content=${encodeURIComponent(content)}&utm_term=${encodeURIComponent(term)}`;

//   Office.context.mailbox.item.body.getAsync("html", function(res) {
//     if (res.status === Office.AsyncResultStatus.Succeeded) {
//       let body = res.value;
//       console.log(body);
//       // Regex to match all <a href="..."> links
//       body = body.replace(/<a\b[^>]*\bhref=["']?([^"'>]+)["']?[^>]*>/gi, (match, url) => {
//         try {
//           let newUrl = new URL(url, "https://dummybase.com"); // dummy base for relative URLs
//           newUrl.search = utm; // append UTM
//           return match.replace(url, newUrl.toString().replace("https://dummybase.com", ""));
//         } catch (e) {
//           return match; // skip invalid URLs
//         }
//       });

//       Office.context.mailbox.item.body.setAsync(body, { coercionType: Office.CoercionType.Html }, function(setRes) {
//         if (setRes.status === Office.AsyncResultStatus.Succeeded) {
//           console.log("✅ All links in signature updated with UTM!");
//         } else {
//           console.error("❌ Failed updating links:", setRes.error);
//         }
//       });

//     } else {
//       console.error("❌ Could not fetch email body:", res.error);
//     }
//   });
// }

// ✅ Final: Update UTM inside Signature Only
async function updateUTMInSignature() {
  // 🎯 Step 3: Skip processing for meeting invites
  if (
    Office.context.mailbox.item.itemType ===
    Office.MailboxEnums.ItemType.Appointment
  ) {
    console.log("📅 Detected meeting invite — skipping UTM update.");
    return;
  }

  // 🆕 Always refresh UTM fields before updating
  await populateFields();

  const campaign = document.getElementById("utm_campaign").value || "";
  const source = document.getElementById("utm_source").value || "";
  const medium = document.getElementById("utm_medium").value || "";
  const content = document.getElementById("utm_content").value || "";
  const term = document.getElementById("utm_term").value || "";

  Office.context.mailbox.item.body.getAsync("html", function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("❌ Could not fetch email body:", res.error);
      return;
    }

    let body = res.value;

    // 🎯 Step 1: Locate the signature block
    const sigMatch = body.match(
      /<div[^>]*(id|class)=["'][^"']*custom-signature[^"']*["'][^>]*>([\s\S]*?)<\/div>/i
    );
    if (!sigMatch) {
      console.log("⚠️ No signature block found — skipping UTM update.");
      return;
    }

    let signatureHtml = sigMatch[2];

    // 🎯 Step 2: Check if signature has any <a> links
    if (!/<a\b[^>]*href=/i.test(signatureHtml)) {
      console.log("ℹ️ No URLs found in signature — skipping UTM update.");
      return;
    }

    // 🎯 Step 3: Replace only links inside the signature
    signatureHtml = signatureHtml.replace(
      /<a\b[^>]*href=["']?([^"'>]+)["']?[^>]*>/gi,
      (match, url) => {
        try {
          if (url.startsWith("mailto:") || url.startsWith("tel:")) return match;

          const lowerUrl = url.toLowerCase();
          const meetingDomains = [
            "teams.microsoft.com",
            "zoom.us",
            "meet.google.com",
            "outlook.office.com/calendar",
            "webex.com",
            "gotomeeting.com",
            "meet.jit.si",
          ];

          if (
            meetingDomains.some((domain) => lowerUrl.includes(domain)) ||
            lowerUrl.includes("meeting")
          ) {
            console.log("📅 Skipping meeting link:", url);
            return match;
          }

          // ✅ Step 4: Preserve existing parameters
          const newUrl = new URL(url, "https://dummybase.com");
          const params = new URLSearchParams(newUrl.search);

          params.set("utm_campaign", campaign);
          params.set("utm_source", source);
          params.set("utm_medium", medium);
          params.set("utm_content", content);
          params.set("utm_term", term);

          newUrl.search = params.toString();

          return match.replace(
            url,
            newUrl.toString().replace("https://dummybase.com", "")
          );
        } catch (e) {
          console.warn("⚠️ Invalid or non-HTTP URL skipped:", url);
          return match;
        }
      }
    );

    // 🎯 Step 5: Put updated signature back into body
    body = body.replace(
      sigMatch[0],
      `<div id="custom-signature">${signatureHtml}</div>`
    );

    Office.context.mailbox.item.body.setAsync(
      body,
      { coercionType: Office.CoercionType.Html },
      function (setRes) {
        if (setRes.status === Office.AsyncResultStatus.Succeeded) {
          console.log(
            "✅ UTM updated successfully — only signature links modified."
          );
        } else {
          console.error("❌ Failed to update signature links:", setRes.error);
        }
      }
    );
  });
}

let quill;

// Open modal + init Quill
document.getElementById("openEditorBtn").onclick = () => {
  document.getElementById("editorModal").style.display = "block";

  if (!quill) {
    quill = new Quill("#quillEditor", {
      theme: "snow",
      modules: {
        toolbar: [
          ["bold", "italic", "underline", "strike"],
          [{ font: [] }, { size: [] }],
          [{ color: [] }, { background: [] }],
          [{ align: [] }],
          ["link", "image"],
          [{ list: "ordered" }, { list: "bullet" }],
        ],
      },
    });

    // Load saved signature into editor
    OfficeRuntime.storage
      .getItem("customSignature")
      .then((savedSig) => {
        if (savedSig) {
          quill.root.innerHTML = savedSig;
        }
      })
      .catch((err) => console.error("⚠️ Failed to load saved signature:", err));
  }
};

// Close modal
function closeEditor() {
  document.getElementById("editorModal").style.display = "none";
}

// // Save Signature (only in localStorage, not in body)
// document.getElementById("saveSignatureBtn").onclick = () => {
//   const sigHTML = quill.root.innerHTML;
//   localStorage.setItem("customSignature", sigHTML);

//   console.log("✅ Signature saved in localStorage");

//   // Close modal after save
//   document.getElementById("editorModal").style.display = "none";

//   // Insert signature in email
//   insertSignature();
// };

// ✅ Save Signature using OfficeRuntime.storage (cross-platform safe)
document.getElementById("saveSignatureBtn").onclick = async () => {
  const sigHTML = quill.root.innerHTML;

  try {
    await OfficeRuntime.storage.setItem("customSignature", sigHTML);
    console.log("✅ Signature saved in OfficeRuntime storage");
  } catch (e) {
    console.error("❌ Failed to save signature:", e);
  }

  document.getElementById("editorModal").style.display = "none";
  insertSignature();
};

// Insert signature into email body (replace if exists)
// function insertSignature() {
//   const sigHTML = await OfficeRuntime.storage.getItem("customSignature") || "";

//   if (!sigHTML) {
//     console.log("⚠️ No signature saved yet!");
//     return;
//   }

//   Office.context.mailbox.item.body.getAsync("html", (res) => {
//     if (res.status === Office.AsyncResultStatus.Succeeded) {
//       let currentBody = res.value;

//       // Remove ALL old signatures
//       currentBody = currentBody.replace(
//         /<div[^>]*id="[^"]*custom-signature"[^>]*>[\s\S]*?<\/div>/gi,
//         ""
//       );

//       // Add fresh one at the end
//       const newBody =
//         currentBody + `<div id="custom-signature">${sigHTML}</div>`;

//       Office.context.mailbox.item.body.setAsync(
//         newBody,
//         { coercionType: Office.CoercionType.Html },
//         (res2) => {
//           if (res2.status === Office.AsyncResultStatus.Succeeded) {
//             console.log("✅ Signature replaced successfully!");
//           } else {
//             console.error("❌ Error inserting signature:", res2.error);
//           }
//         }
//       );
//     }
//   });
// }

async function insertSignature() {
  const sigHTML =
    (await OfficeRuntime.storage.getItem("customSignature")) || "";
  if (!sigHTML) {
    console.log("⚠️ No signature saved yet!");
    return;
  }

  Office.context.mailbox.item.body.getAsync("html", (res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      let currentBody = res.value;

      currentBody = currentBody.replace(
        /<div[^>]*id="[^"]*custom-signature"[^>]*>[\s\S]*?<\/div>/gi,
        ""
      );

      const newBody =
        currentBody + `<div id="custom-signature">${sigHTML}</div>`;

      Office.context.mailbox.item.body.setAsync(
        newBody,
        { coercionType: Office.CoercionType.Html },
        (res2) => {
          if (res2.status === Office.AsyncResultStatus.Succeeded) {
            console.log("✅ Signature replaced successfully!");
          } else {
            console.error("❌ Error inserting signature:", res2.error);
          }
        }
      );
    }
  });
}
