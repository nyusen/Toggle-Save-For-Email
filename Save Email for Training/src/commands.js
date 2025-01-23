/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var mailboxItem;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
function validateBody(event) {
    // Get the body of the message
    Office.context.ui.displayDialogAsync('https://localhost:3000/dialog.html', 
      {height: 30, width: 20, displayInIframe: true},
      function (result) {
          if (result.status === Office.AsyncResultStatus.Failed) {
              // Handle error
              console.error(result.error.message);
              event.completed({ allowEvent: false });
              return;
          } else {
              // Get dialog instance
              var dialog = result.value;
              
              // Add event handler for dialog events
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                  dialog.close();
                  if (arg.message === 'yes') {
                      // User clicked Yes, allow send
                      event.completed({ allowEvent: true });
                  } else {
                      // User clicked No or closed dialog, block send
                      event.completed({ allowEvent: false });
                  }
                  return;
              });
          }
      }
  );
}