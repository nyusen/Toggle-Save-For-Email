/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const config = {
    clientId: "6d7e781e-9cf5-48ff-8c05-b697ca1a90e3",
    redirectUri: "https://nyusen.github.io/Save-Email-For-Training/auth-callback.html",
    authEndpoint: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    tokenEndpoint: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    scopes: "openid profile email Mail.Read Mail.Send",
};

let accessToken = null;
let idToken = null;

// Make functions globally available
window.validateBody = validateBody;
window.saveForTraining = saveForTraining;

Office.onReady(async (info) => {
    if (info.host) {
        // Check if we have tokens in session storage
        accessToken = sessionStorage.getItem('accessToken');
        idToken = sessionStorage.getItem('idToken');
        
        if (accessToken) {
            updateUI(true);
        } else {
            updateUI(false);
        }

        // Add click handler for sign-in button
        document.getElementById('signInButton').onclick = handleSignIn;
    }
});

// Function to handle sign-in button click
function handleSignIn() {
    console.log("Handle sign in called");
    return new Promise((resolve, reject) => {
        // Generate random state and PKCE verifier
        const state = generateRandomString(32);
        const codeVerifier = generateRandomString(64);
        const nonce = generateRandomString(16);
        
        // Store state and code verifier in parent window's session storage
        window.sessionStorage.setItem('authState', state);
        window.sessionStorage.setItem('codeVerifier', codeVerifier);

        console.log('Generated values:', {
            state: state,
            codeVerifier: '[GENERATED]',
            nonce: nonce
        });

        // Generate code challenge
        generateCodeChallenge(codeVerifier).then(codeChallenge => {
            // Build the URL to our local sign-in start page
            const startUrl = new URL('https://nyusen.github.io/Save-Email-For-Training/sign-in-start.html');
            startUrl.searchParams.append('state', state);
            startUrl.searchParams.append('nonce', nonce);
            startUrl.searchParams.append('code_challenge', codeChallenge);
            startUrl.searchParams.append('code_challenge_method', 'S256');
            startUrl.searchParams.append('code_verifier', codeVerifier);
            
            // Add delay before opening the auth dialog
            setTimeout(() => {
                console.log("Opening auth dialog: " + startUrl.toString());
                // Open our local page in a dialog, which will then redirect to Microsoft
                Office.context.ui.displayDialogAsync(startUrl.toString(), 
                    {height: 60, width: 30}, 
                    function(result) {
                        if (result.status === Office.AsyncResultStatus.Failed) {
                            console.error(result.error.message);
                            reject(new Error(result.error.message));
                            return;
                        }

                        const dialog = result.value;
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                            console.log("Received message from auth dialog: " + arg.message);
                            try {
                                const message = JSON.parse(arg.message);
                                if (message.type === 'token') {
                                    // Store both tokens
                                    accessToken = message.token;
                                    idToken = message.idToken;
                                    sessionStorage.setItem('accessToken', message.token);
                                    sessionStorage.setItem('idToken', message.idToken);
                                    sessionStorage.setItem('tokenType', message.tokenType);
                                    sessionStorage.setItem('expiresIn', message.expiresIn);
                                    dialog.close();
                                    resolve(message.token);
                                } else if (message.type === 'error') {
                                    dialog.close();
                                    reject(new Error(message.error));
                                }
                            } catch (e) {
                                console.error('Error parsing message:', e);
                                dialog.close();
                                reject(e);
                            }
                        });

                        dialog.addEventHandler(Office.EventType.DialogEventReceived, function(arg) {
                            dialog.close();
                            reject(new Error('Dialog closed'));
                        });
                    }
                );
            }, 1000); // Wait 1 second to ensure previous dialog is fully closed
        }).catch(error => {
            reject(error);
        });
    });
}

// Function to generate random string for PKCE and state
function generateRandomString(length) {
    const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~';
    let text = '';
    for (let i = 0; i < length; i++) {
        text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
}

// Function to generate code challenge from verifier
async function generateCodeChallenge(codeVerifier) {
    const encoder = new TextEncoder();
    const data = encoder.encode(codeVerifier);
    const digest = await window.crypto.subtle.digest('SHA-256', data);
    return base64URLEncode(digest);
}

// Base64URL encoding function
function base64URLEncode(buffer) {
    return btoa(String.fromCharCode(...new Uint8Array(buffer)))
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/, '');
}

// Function to update UI based on sign-in state
function updateUI(isSignedIn) {
    const signInButton = document.getElementById('signInButton');
    const statusMessage = document.getElementById('statusMessage');
    if (isSignedIn) {
        signInButton.textContent = "Signed In";
        signInButton.disabled = true;
        statusMessage.textContent = "You're signed in and ready to use the add-in.";
    } else {
        signInButton.textContent = "Sign in with Microsoft";
        signInButton.disabled = false;
        statusMessage.textContent = "Please sign in to use the add-in.";
    }
    console.log(signInButton)
    console.log(statusMessage)
}

// Function to make authenticated requests to your server
async function makeAuthenticatedRequest(url, options = {}) {
    if (!accessToken) {
        throw new Error('Not authenticated');
    }

    const headers = {
        'Authorization': `Bearer ${idToken}`,
        'Content-Type': 'application/json',
        ...options.headers
    };

    const response = await fetch(url, {
        ...options,
        headers
    });

    if (!response.ok) {
        if (response.status === 401) {
            // Token might be expired, clear it
            accessToken = null;
            sessionStorage.removeItem('accessToken');
            updateUI(false);
            throw new Error('Authentication expired. Please sign in again.');
        }
        throw new Error(`Request failed: ${response.statusText}`);
    }

    return response.json();
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
async function validateBody(event) {
    try {
        // Show the initial save dialog first
        Office.context.ui.displayDialogAsync('https://nyusen.github.io/Save-Email-For-Training/dialog.html', 
            {height: 30, width: 20, displayInIframe: true},
            async function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error(result.error.message);
                    event.completed({ allowEvent: false });
                    return;
                }
                
                const dialog = result.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, async function (arg) {
                    dialog.close();
                    if (arg.message === 'Send and Save') {
                        // Check if user is signed in before saving
                        if (!accessToken) {
                            // Wait for the first dialog to fully close
                            setTimeout(() => {
                                console.log("Showing sign-in dialog");
                                Office.context.ui.displayDialogAsync('https://nyusen.github.io/Save-Email-For-Training/signin-dialog.html', 
                                    {height: 40, width: 30, displayInIframe: true},
                                    function (signInResult) {
                                        if (signInResult.status === Office.AsyncResultStatus.Failed) {
                                            console.error(signInResult.error.message);
                                            event.completed({ allowEvent: false });
                                            return;
                                        }

                                        console.log("Listening for message");

                                        const signInDialog = signInResult.value;
                                        signInDialog.addEventHandler(Office.EventType.DialogMessageReceived, async function (signInArg) {
                                            if (signInArg.message === 'signin') {
                                                // User wants to sign in
                                                signInDialog.close();
                                                // Start the sign-in process
                                                handleSignIn().then((token) => {
                                                    // User is signed in, proceed with saving
                                                    saveForTraining(event);
                                                }).catch((error) => {
                                                    console.error('Error signing in:', error);
                                                    event.completed({ allowEvent: false });
                                                });
                                            } else if (signInArg.message === 'cancel') {
                                                // User doesn't want to sign in
                                                signInDialog.close();
                                                event.completed({ allowEvent: false });
                                            }
                                        });
                                    }
                                );
                            }, 500); // Wait 500ms for the first dialog to close
                        } else {
                            // User is already signed in, proceed with saving
                            await saveForTraining(event);
                        }
                    } else {
                        // User clicked Send Only
                        event.completed({ allowEvent: true });
                    }
                });
            }
        );
    } catch (error) {
        console.error(error);
        event.completed({ allowEvent: false });
    }
}

async function saveForTraining(event) {
    console.log("Saving for Training");
    try {
        const item = Office.context.mailbox.item;
        
        // Get email data using promises
        const [subject, body, sender, recipients] = await Promise.all([
            new Promise((resolve) => item.subject.getAsync(resolve)),
            new Promise((resolve) => item.body.getAsync('text', resolve)),
            new Promise((resolve) => item.from.getAsync(resolve)),
            new Promise((resolve) => item.to.getAsync(resolve))
        ]);
        
        const emailData = {
            subject: subject.value,
            body: body.value,
            sender: sender.value.emailAddress,
            recipients: recipients.value.map(r => r.emailAddress),
            timestamp: new Date().toISOString()
        };

        // Make authenticated request to your server
        await makeAuthenticatedRequest('https://ml-inf-svc-dev.eventellect.com/corpus-collector/api/save-email', {
            method: 'POST',
            body: JSON.stringify(emailData)
        });

        event.completed({ allowEvent: true });
    } catch (error) {
        console.error('Error:', error);
        // Show retry dialog with error message
        const errorMessage = encodeURIComponent(error.message || 'An unknown error occurred');
        Office.context.ui.displayDialogAsync(`https://nyusen.github.io/Save-Email-For-Training/retry-dialog.html?error=${errorMessage}`, 
            {height: 30, width: 20, displayInIframe: true},
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error(asyncResult.error.message);
                    event.completed({ allowEvent: false });
                    return;
                }
                
                const dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                    dialog.close();
                    if (arg.message === 'retry') {
                        saveForTraining(event);
                    } else {
                        event.completed({ allowEvent: true });
                    }
                });
            }
        );
    }
}