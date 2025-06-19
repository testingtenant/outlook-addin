Office.onReady(() => {
  // Listen for messages from the parent (launchevent.js)
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    (arg) => {
      const message = JSON.parse(arg.message);
      const recipientList = document.getElementById("recipient-list");
      message.externalRecipients.forEach((recipient) => {
        const li = document.createElement("li");
        li.textContent = recipient;
        recipientList.appendChild(li);
      });
    }
  );

  // Handle Send Anyway button click
  document.getElementById("sendButton").addEventListener("click", () => {
    Office.context.ui.messageParent("send");
  });

  // Handle Don't Send button click
  document.getElementById("cancelButton").addEventListener("click", () => {
    Office.context.ui.messageParent("cancel");
  });
});