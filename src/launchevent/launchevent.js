// Define the customer domain
const customerDomain = "@bizwind.co.jp";

// Handler for the OnMessageSend event
async function onMessageSendHandler(event) {
  try {
    let externalRecipients = [];

    // Function to check a single email address
    function checkEmail(email, field) {
      let cleanedEmail = email;
      const match = cleanedEmail.match(/<(.+?)>|[^<>\s]+/);
      cleanedEmail = match ? match[1] || match[0] : cleanedEmail;
      cleanedEmail = cleanedEmail.trim().toLowerCase();
      const domain = customerDomain.toLowerCase();

      console.log(`Checking ${field} email: ${cleanedEmail}`);
      console.log(`Ends with ${domain}? ${cleanedEmail.endsWith(domain)}`);

      if (!cleanedEmail.endsWith(domain)) {
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
      console.log("Failed to get To recipients.");
    }

    // Process CC recipients
    if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
      ccResult.value.forEach((recipient) => checkEmail(recipient.emailAddress, "CC"));
    } else {
      console.log("Failed to get CC recipients.");
    }

    // Process BCC recipients
    if (bccResult.status === Office.AsyncResultStatus.Succeeded) {
      bccResult.value.forEach((recipient) => checkEmail(recipient.emailAddress, "BCC"));
    } else {
      console.log("Failed to get BCC recipients.");
    }

    // Decide whether to show popup
    if (externalRecipients.length > 0) {
      console.log(`External recipients found: ${externalRecipients.join(", ")}`);
      event.completed({
        allowEvent: false,
        errorMessage:
          "You are sending this email to external recipients:\n\n" +
          externalRecipients.join("\n") +
          "\n\nAre you sure you want to send it?",
      });
    } else {
      console.log("No external recipients found, allowing send.");
      event.completed({ allowEvent: true });
    }
  } catch (error) {
    console.log("Error in onMessageSendHandler: " + error);
    event.completed({ allowEvent: true }); // Fallback to allow send on error
  }
}

// Ensure Office API is ready before associating the event handler
Office.onReady((info) => {
  if (info.platform === Office.PlatformType.PC || info.platform == null) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }
}).catch((error) => {
  console.log("Error initializing Office API: " + error);
});
