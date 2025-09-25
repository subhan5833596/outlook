/* global Office, OfficeRuntime */

(function () {
    // Existing helper functions
    async function getDefault(event) {
        try {
            Office.context.mailbox.item.body.getAsync(
                Office.CoercionType.Html,
                { asyncContext: event },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("Body (HTML):", asyncResult.value);
                    } else {
                        console.error("Error getting body:", asyncResult.error);
                    }
                    event.completed();
                }
            );
        } catch (e) {
            console.error("getDefault failed:", e);
            event.completed();
        }
    }

    async function setDefaultData(event) {
        try {
            Office.context.mailbox.item.body.setAsync(
                "<p>UTM Parameters applied successfully ✅</p>",
                { coercionType: Office.CoercionType.Html },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("Body updated");
                    } else {
                        console.error("Error setting body:", asyncResult.error);
                    }
                    event.completed();
                }
            );
        } catch (e) {
            console.error("setDefaultData failed:", e);
            event.completed();
        }
    }

    async function validateBody(event) {
        try {
            Office.context.mailbox.item.body.getAsync(
                Office.CoercionType.Html,
                { asyncContext: event },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        const body = asyncResult.value || "";
                        const hasUTM = body.includes("utm_");
                        console.log("validateBody:", hasUTM ? "✅ UTM found" : "⚠️ No UTM tags");
                    } else {
                        console.error("Error validating body:", asyncResult.error);
                    }
                    event.completed();
                }
            );
        } catch (e) {
            console.error("validateBody failed:", e);
            event.completed();
        }
    }

    // 🔥 New function for signature editing
    async function editSignature(event) {
        try {
            Office.context.mailbox.item.setSignatureAsync(
                "<p>--<br>Best Regards,<br><strong>Your UTM Signature</strong></p>",
                { coercionType: Office.CoercionType.Html },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("Signature set successfully");
                    } else {
                        console.error("Error setting signature:", asyncResult.error);
                    }
                    event.completed();
                }
            );
        } catch (e) {
            console.error("editSignature failed:", e);
            event.completed();
        }
    }

    // Register actions with Office.js
    Office.actions.associate("GetDefault", getDefault);
    Office.actions.associate("SetDefaultData", setDefaultData);
    Office.actions.associate("validateBody", validateBody);

    // 👇 Newly added registration for signature
    Office.actions.associate("EditSignature", editSignature);
})();
