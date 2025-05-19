Office.onReady(function () {
    if (!Office.context.mailbox) {
        console.error("This add-in only works in Outlook.");
        return;
    }

    Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemSend, checkRecipients);
});

function checkRecipients(eventArgs) {
    const organizationDomain = "4h72mt.onmicrosoft.com";  
    let warn = false;

    // Check the 'to' recipients
    Office.context.mailbox.item.to.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const recipients = asyncResult.value;
            if (recipients.length > 0) {
                recipients.forEach((email) => {
                    if (email.emailAddress.indexOf("@") > 0) {
                        let domain = email.emailAddress.split('@')[1].toLowerCase();
                        if (domain !== organizationDomain) {
                            warn = true;
                        }
                    }
                });
            }

            if (warn) {
                showPopup();
                eventArgs.completed({ allowEvent: false });
            } else {
                eventArgs.completed({ allowEvent: true });
            }
        } else {
            console.error("Failed to retrieve recipients:", asyncResult.error);
            eventArgs.completed({ allowEvent: true });
        }
    });
}

function showPopup() {
    const popup = document.getElementById("warningPopup");
    popup.style.display = "flex";
    document.getElementById("continueBtn").onclick = () => {
        popup.style.display = "none";
    };
    document.getElementById("cancelBtn").onclick = () => {
        popup.style.display = "none";
        Office.context.mailbox.item.cancelSendAsync();
    };
}
