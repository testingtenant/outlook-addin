// Handler for the OnMessageSend event
async function onMessageSendHandler(event) {
  try {
    let externalRecipients = [];

    // Load internal domains from roamingSettings or fallback to user's email domain
    let internalDomains = Office.context.roamingSettings.get("internalDomains") || [];
    if (internalDomains.length === 0) {
      const userEmail = Office.context.mailbox.userProfile.emailAddress;
      if (!userEmail.includes('@')) {
        console.log('Invalid user email format, allowing send.');
        event.completed({ allowEvent: true });
        return;
      }
      internalDomains = [userEmail.substring(userEmail.lastIndexOf('@')).toLowerCase()];
    }
    console.log(`Internal domains: ${internalDomains.join(', ')}`);

    // Function to check a single email address
    function checkEmail(email, field) {
      if (!email) return;
      let cleanedEmail = email.trim().toLowerCase();
      const match = cleanedEmail.match(/<(.+?)>|[^<>\s]+/);
      if (!match) {
        console.log(`Invalid email format in ${field}: ${email}`);
        return;
      }
      cleanedEmail = match[1] || match[0];
      if (!cleanedEmail.includes('@')) {
        console.log(`Skipping invalid email in ${field}: ${cleanedEmail}`);
        return;
      }
      const isExternal = !internalDomains.some(domain => cleanedEmail.endsWith(domain));
      if (isExternal) {
        externalRecipients.push(`${field}: ${cleanedEmail}`);
      }
    }

    // Parallelize the retrieval of To, CC, and BCC recipients
    const [toResult, ccResult, bccResult] = await Promise.all([
      new Promise((resolve) => Office.context.mailbox.item.to.getAsync(resolve)),
      new Promise((resolve) => Office.context.mailbox.item.cc.getAsync(resolve)),
      new Promise((resolve) => Office.context.mailbox.item.bcc.getAsync(resolve)),
    ]);

    // Process To recipients
    if (toResult.status === Office.AsyncResultStatus.Succeeded) {
      toResult.value.forEach((recipient) => checkEmail(recipient.emailAddress, "To"));
    } else {
      console.log(`Failed to get To recipients: ${toResult.error.message}`);
    }

    // Process CC recipients
    if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
      ccResult.value.forEach((recipient) => checkEmail(recipient.emailAddress, "CC"));
    } else {
      console.log(`Failed to get CC recipients: ${ccResult.error.message}`);
    }

    // Process BCC recipients
    if (bccResult.status === Office.AsyncResultStatus.Succeeded) {
      bccResult.value.forEach((recipient) => checkEmail(recipient.emailAddress, "BCC"));
    } else {
      console.log(`Failed to get BCC recipients: ${bccResult.error.message}`);
    }

    // Decide whether to show popup
    if (externalRecipients.length > 0) {
      console.log(`External recipients found: ${externalRecipients.join(", ")}`);
      // Use custom dialog instead of errorMessage
      const dialogUrl = "https://testingtenant.github.io/outlook-addin/src/dialog/dialog.html";
      Office.context.ui.displayDialogAsync(dialogUrl, { height: 40, width: 30 }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(`Failed to open dialog: ${asyncResult.error.message}`);
          event.completed({ allowEvent: true }); // Fallback to allow send
          return;
        }
        const dialog = asyncResult.value;
        dialog.args = { recipients: externalRecipients }; // Pass recipients to dialog
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          if (arg.message && typeof arg.message === "string") {
            try {
              const message = JSON.parse(arg.message);
              if (message.action === "send") {
                event.completed({ allowEvent: true });
              } else {
                event.completed({ allowEvent: false });
              }
            } catch (e) {
              console.log(`Error parsing dialog message: ${e.message}`);
              event.completed({ allowEvent: false });
            }
          } else {
            console.log("Invalid dialog message received");
            event.completed({ allowEvent: false });
          }
        });
      });
    } else {
      console.log("No external recipients found, allowing send.");
      event.completed({ allowEvent: true });
    }
  } catch (error) {
    console.log(`Error in onMessageSendHandler: ${error.message}`);
    event.completed({ allowEvent: true }); // Fallback to allow send
  }
}

// Ensure Office API is ready before associating the event handler
Office.onReady((info) => {
  if (info.platform === Office.PlatformType.PC || info.platform == null) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }
}).catch((error) => {
  console.log(`Error initializing Office API: ${error.message}`);
});
