function addBccAddress(event) {
  Office.context.mailbox.item.bcc.addAsync(
    { emailAddress: "emailtosalesforce@2lb252qj41tlq82ooamrh5mvk5uidil4qtbu88surrqi7brj1y.8-a0kdeay.na104.le.salesforce.com" },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(asyncResult.error.message);
      } else {
        console.log("BCC address added successfully.");
      }
      event.completed();
    }
  );
}
