Office.onReady(() => {
    if (!Office.context.mailbox) {
        console.warn("This add-in only works in Outlook.");
        document.getElementById("message").innerText = "This add-in only works in Outlook.";
        return;
    }

    // Only run if this is an email message read form
    if (Office.context.mailbox.item && Office.context.mailbox.item.to) {
        Office.context.mailbox.item.to.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const recipients = asyncResult.value;
                const externalRecipients = recipients.filter(recipient => {
                    return !recipient.emailAddress.endsWith("@4h72mt.onmicrosoft.com") && !recipient.emailAddress.endsWith("@bizwind.co.jp");
                });

                if (externalRecipients.length > 0) {
                    const warningMessage = "Warning: You are sending an email to an external recipient.";
                    document.getElementById("message").innerText = warningMessage;
                    console.warn(warningMessage);
                } else {
                    document.getElementById("message").innerText = "No external recipients detected.";
                }
            } else {
                console.error("Failed to get recipient information.", asyncResult.error.message);
                document.getElementById("message").innerText = "Failed to retrieve recipient information.";
            }
        });
    } else {
        document.getElementById("message").innerText = "Cannot get recipient information for this item.";
    }
});
