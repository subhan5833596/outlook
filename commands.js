/* global Office */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    console.log("UTM Manager Add-in loaded successfully.");
  }
});

// Associate your functions with Office actions
Office.actions.associate("SetDefaultData", SetDefaultData);
Office.actions.associate("validateBody", validateBody);
Office.actions.associate("GetDefault", GetDefault);
Office.actions.associate("EditSignature", EditSignature);

// -------------------- Functions --------------------

function SetDefaultData(event) {
  console.log("SetDefaultData called");
  event.completed();
}

function validateBody(event) {
  console.log("validateBody called");
  event.completed({ allowEvent: true });
}

function GetDefault(event) {
  console.log("GetDefault called");
  event.completed();
}

// 🆕 EditSignature Function
function EditSignature(event) {
  try {
    const testSignature = "<p>---<br><b>UTM Manager Test Signature</b></p>";

    Office.context.mailbox.item.body.setSignatureAsync(
      testSignature,
      { coercionType: "html" },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("✅ Signature inserted successfully.");
        } else {
          console.error("❌ Error inserting signature:", asyncResult.error);
        }
        event.completed(); // Always complete the event
      }
    );
  } catch (err) {
    console.error("❌ EditSignature exception:", err);
    event.completed();
  }
}
