/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { createNestablePublicClientApplication } from "@azure/msal-browser";

let pca = undefined;
Office.onReady(async (info) => {
  if (info.host) {
    // Initialize the public client application
    pca = await createNestablePublicClientApplication({
      auth: {
        clientId: "6d7e781e-9cf5-48ff-8c05-b697ca1a90e3",
        authority: "https://login.microsoftonline.com/common"
      },
    });
  }
});

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
async function validateBody(event) {
    Office.context.ui.displayDialogAsync('https://ml-inf-svc-dev.eventellect.com/outlook-save-for-training/dialog.html', 
      {height: 30, width: 20, displayInIframe: true},
      async function (result) {
          if (result.status === Office.AsyncResultStatus.Failed) {
              // Handle error
              console.error(result.error.message);
              event.completed({ allowEvent: false });
              return;
          } else {
              // Get dialog instance
              var dialog = result.value;
              
              // Add event handler for dialog events
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, async function (arg) {
                  dialog.close();
                  if (arg.message === 'Send and Save') {
                      // User clicked Send and Save (for training)
                      await saveForTraining(event);
                  } else {
                      // User clicked Send Only
                      event.completed({ allowEvent: true });
                  }
              });
          }
      }
  );
}

async function saveForTraining(event) {
    // Get email data from the current item
    const tokenRequest = {
        scopes: ["Mail.Read", "Mail.Send", "openid", "profile", "email"],
    }
    let accessToken = null;

    try {
        console.log("Trying to acquire token silently...");
        const userAccount = await pca.acquireTokenSilent(tokenRequest);
        console.log("Acquired token silently.");
        accessToken = userAccount.accessToken;
    } catch (error) {
        console.log(`Unable to acquire token silently: ${error}`);
    }

    if (accessToken === null) {
        // Acquire token silent failure. Send an interactive request via popup.
        try {
          console.log("Trying to acquire token interactively...");
          const userAccount = await pca.acquireTokenPopup(tokenRequest);
          console.log("Acquired token interactively.");
          accessToken = userAccount.accessToken;
        } catch (popupError) {
          // Acquire token interactive failure.
          console.log(`Unable to acquire token interactively: ${popupError}`);
        }
    }

    // Log error if both silent and popup requests failed.
    if (accessToken === null) {
        console.error(`Unable to acquire access token.`);
        return;
    }
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
                                        fetch('https://localhost:8000/save-email', {
                                            method: 'POST',
                                            headers: {
                                                'Content-Type': 'application/json',
                                                'Authorization': `Bearer ${identityToken}`
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
                                }
                            )
                        }
                    )
                }
            )
        }
    )
}