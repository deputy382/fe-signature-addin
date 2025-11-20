
Office.onReady(() => {
  document.getElementById("showMessage").onclick = () => {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
      type: "informationalMessage",
      message: "Hello from your task pane!",
      icon: "icon16",
      persistent: false
    });
  };
});
