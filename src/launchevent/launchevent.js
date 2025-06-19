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

    // Decide whether to show styled notification
    if (externalRecipients.length > 0) {
      console.log(`External recipients found: ${externalRecipients.join(", ")}`);

      const htmlList = externalRecipients.map(item => `<span style="color:red;">${item}</span>`).join("<br>");

      Office.context.mailbox.item.notificationMessages.addAsync("externalWarning", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `<b>External Recipients Detected:</b><br>${htmlList}`,
        icon: "Icon.80x80",
        persistent: true
      }, () => {
        event.completed({
          allowEvent: false,
          errorMessage: "You're sending this email to external recipients. Please check the red-highlighted recipients above."
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
