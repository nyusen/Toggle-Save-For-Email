/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// var mailboxItem;

Office.initialize = function (reason) {
    // mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
function validateBody(event) {
    Office.context.ui.displayDialogAsync('https://ml-inf-svc-dev.eventellect.com/outlook-save-for-training/dialog.html', 
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
                  if (arg.message === 'Send and Save') {
                      // User clicked Send and Save (for training)
                      saveForTraining(event);
                  } else {
                      // User clicked Send Only
                      event.completed({ allowEvent: true });
                  }
              });
          }
      }
  );
}

function saveForTraining(event) {
    // Get email data from the current item
    const item = Office.context.mailbox.item;

    item.subject.getAsync(
        {asyncContext: {currentItem: item}},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
                event.completed({ allowEvent: false });
                return;
            }
            const subject = result.value;
            item.body.getAsync(
                Office.CoercionType.Text,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.error(result.error.message);
                        event.completed({ allowEvent: false });
                        return;
                    }
                    const body = result.value;
                    item.from.getAsync(
                        {asyncContext: {currentItem: item}},
                        function (result) {
                            if (result.status === Office.AsyncResultStatus.Failed) {
                                console.error(result.error.message);
                                event.completed({ allowEvent: false });
                                return;
                            }
                            const sender = result.value;
                            item.to.getAsync(
                                {asyncContext: {currentItem: item}},
                                function (result) {
                                    if (result.status === Office.AsyncResultStatus.Failed) {
                                        console.error(result.error.message);
                                        event.completed({ allowEvent: false });
                                        return;
                                    }
                                    const recipients = result.value;
                                    const timestamp = new Date().toISOString();
                                    const emailData = {
                                        subject,
                                        body,
                                        sender: sender.emailAddress,
                                        recipients: recipients.map(r => r.emailAddress),
                                        timestamp
                                    };
                                    function attemptSave() {
                                        fetch('https://ml-inf-svc-dev.eventellect.com/corpus-collector/api/save-email', {
                                            method: 'POST',
                                            headers: {
                                                'Content-Type': 'application/json',
                                            },
                                            body: JSON.stringify(emailData)
                                        })
                                        .then(response => {
                                            if (!response.ok) {
                                                throw new Error('Network response was not ok');
                                            }
                                            return response.json();
                                        })
                                        .then(data => {
                                            console.log('Success:', data);
                                            if (event) {
                                                event.completed({ allowEvent: true });
                                            }
                                        })
                                        .catch((error) => {
                                            console.error('Error:', error);
                                            // Show retry dialog
                                            Office.context.ui.displayDialogAsync('https://ml-inf-svc-dev.eventellect.com/outlook-save-for-training/retry-dialog.html', 
                                                {height: 30, width: 20, displayInIframe: true},
                                                function (asyncResult) {
                                                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                                        console.error(asyncResult.error.message);
                                                        if (event) {
                                                            event.completed({ allowEvent: false });
                                                        }
                                                    } else {
                                                        // Get dialog instance
                                                        var dialog = asyncResult.value;
                                                        
                                                        // Add event handler for dialog events
                                                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                                                            dialog.close();
                                                            if (arg.message === 'retry') {
                                                                // Try saving again
                                                                attemptSave();
                                                            } else if (arg.message === 'send-only') {
                                                                // Proceed with sending without saving
                                                                if (event) {
                                                                    event.completed({ allowEvent: true });
                                                                }
                                                            }
                                                        });
                                                    }
                                                }
                                            );
                                        });
                                    }
                            
                                    // Initial save attempt
                                    attemptSave();
                                })
                            })
                        })
                    })
                }