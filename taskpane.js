Office.onReady(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemSend, checkRecipients);
});

function checkRecipients(eventArgs) {
    const organizationDomain = "@4h72mt.onmicrosoft.com";  
    let warn = false;
    Office.context.mailbox.item.to.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const recipients = asyncResult.value;
            recipients.forEach((email) => {
                if (email.emailAddress.indexOf("@") > 0) {
                    let domain = email.emailAddress.split('@')[1].toLowerCase();
                    if (domain !== organizationDomain) {
                        warn = true;
                    }
                }
            });
            if (warn) {
                showPopup();
                eventArgs.completed({ allowEvent: false });
            } else {
                eventArgs.completed({ allowEvent: true });
            }
        } else {
            console.error("Failed to retrieve recipients.");
            eventArgs.completed({ allowEvent: true });
        }
    });
}

function showPopup() {
    const popup = document.getElementById("warningPopup");
    popup.style.display = "block";
    document.getElementById("continueBtn").onclick = () => {
        popup.style.display = "none";
    };
    document.getElementById("cancelBtn").onclick = () => {
        popup.style.display = "none";
        Office.context.mailbox.item.cancelSendAsync();
    };
}