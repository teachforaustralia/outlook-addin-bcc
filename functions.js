function addBccAddress(event) {
  Office.context.mailbox.item.bcc.addAsync(
    { emailAddress: "itmanager@teachforaustralia.org" },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(asyncResult.error.message);
      }.log("BCC address added successfully.");
      }
      event.completed();
    }
  );
}
