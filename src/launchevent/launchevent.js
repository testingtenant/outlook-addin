// Utility to add a timeout to a promise
const withTimeout = (promise, ms) => {
  const timeout = new Promise((_, reject) =>
    setTimeout(() => reject(new Error("Operation timed out")), ms)
  );
  return Promise.race([promise, timeout]);
};

// Handler for the OnMessageSend event
async function onMessageSendHandler(event) {
  const handlerTimeout = 4000; // 4 seconds to stay within Outlook's ~5s limit
  try {
    // Wrap the entire operation in a timeout
    const result = await withTimeout(
      (async () => {
        let externalRecipients = [];
        let externalDomains = new Set();

        // Load internal domains from roamingSettings or fallback to user's email domain
        let internalDomains = Office.context.roamingSettings.get("internalDomains") || [];
        if (internalDomains.length === 0) {
          const userEmail = Office.context.mailbox.userProfile.emailAddress;
          if (!userEmail || !userEmail.includes('@')) {
            console.error("Invalid or missing user email, blocking send for safety.");
            return {
              allowEvent: false,
              errorMessage: "Error: Unable to determine internal domain. Please contact support.",
            };
          }
          internalDomains = [userEmail.substring(userEmail.lastIndexOf('@')).toLowerCase()];
        }
        console.log(`Internal domains: ${internalDomains.join(', ')}`);

        // Function to check a single email address
        function checkEmail(email, field) {
          if (!email) {
            console.log(`Empty email in ${field}, skipping.`);
            return;
          }
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
            const domain = `@${cleanedEmail.split('@')[1]}`;
            externalDomains.add(domain);
          }
        }

        // Parallelize recipient retrieval with individual timeouts
        const [toResult, ccResult, bccResult] = await Promise.all([
          withTimeout(
            new Promise((resolve) => Office.context.mailbox.item.to.getAsync(resolve)),
            1500
          ),
          withTimeout(
            new Promise((resolve) => Office.context.mailbox.item.cc.getAsync(resolve)),
            1500
          ),
          withTimeout(
            new Promise((resolve) => Office.context.mailbox.item.bcc.getAsync(resolve)),
            1500
          ),
        ]);

        // Process To recipients
        if (toResult.status === Office.AsyncResultStatus.Succeeded) {
          toResult.value.forEach((recipient) => checkEmail(recipient.emailAddress, "To"));
        } else {
          console.error(`Failed to get To recipients: ${toResult.error?.message}`);
        }

        // Process CC recipients
        if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
          ccResult.value.forEach((recipient) => checkEmail(recipient.emailAddress, "CC"));
        } else {
          console.error(`Failed to get CC recipients: ${ccResult.error?.message}`);
        }

        // Process BCC recipients
        if (bccResult.status === Office.AsyncResultStatus.Succeeded) {
          bccResult.value.forEach((recipient) => checkEmail(recipient.emailAddress, "BCC"));
        } else {
          console.error(`Failed to get BCC recipients: ${bccResult.error?.message}`);
        }

        // Decide whether to show popup
        if (externalRecipients.length > 0) {
          console.log(`External recipients found: ${externalRecipients.join(", ")}`);
          console.log(`External domains found: ${Array.from(externalDomains).join(", ")}`);
          const message =
            "You are sending this email to external recipients:\n\n" +
            "__________________________________________________\n\n" +
            "Domain list\n" +
            Array.from(externalDomains)
              .map(domain => `→${domain}`)
              .join("\n") +
            "\n__________________________________________________\n\n" +
            "Email list\n" +
            externalRecipients.join("\n") +
            "\n\nAre you sure you want to send it?";
          return { allowEvent: false, errorMessage: message };
        } else {
          console.log("No external recipients found, allowing send.");
          return { allowEvent: true };
        }
      })(),
      handlerTimeout
    );

    event.completed(result);
  } catch (error) {
    console.error("Error in onMessageSendHandler:", error.message, error.stack);
    // Block send on error to prevent accidental external sends
    event.completed({
      allowEvent: false,
      errorMessage: "Error: Unable to verify recipients. Please try again or contact support.",
    });
  }
}

// Initialize Office API
Office.onReady((info) => {
  if (info.platform === Office.PlatformType.PC || info.platform === null) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }
}).catch((error) => {
  console.error("Error initializing Office API:", error.message, error.stack);
});
