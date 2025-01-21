Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Register the function that will handle the button click
        Office.actions.associate("sendAndSave", sendAndSaveHandler);
    }
});

async function sendAndSaveHandler(event) {
    try {
        // Get the current message
        const item = Office.context.mailbox.item;
        
        // Get email content
        const subject = item.subject;
        const body = await getEmailBody();
        const sender = Office.context.mailbox.userProfile.emailAddress;
        const recipients = getRecipients();
        
        // Create email data object
        const emailData = {
            subject: subject,
            body: body,
            sender: sender,
            recipients: recipients,
            timestamp: new Date().toISOString()
        };

        // Send the email first
        await sendEmail();

        // Save to Azure Blob Storage
        await saveToBlob(emailData);

        // Show success message
        showStatus("Email sent and saved successfully!", "success");
        
        event.completed();
    } catch (error) {
        console.error("Error:", error);
        showStatus("Error: " + error.message, "error");
        event.completed({ allowEvent: false });
    }
}

function getRecipients() {
    const recipients = Office.context.mailbox.item.to;
    return recipients.map(recipient => recipient.emailAddress);
}

async function getEmailBody() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(new Error(result.error.message));
            }
        });
    });
}

async function sendEmail() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.send((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
            } else {
                reject(new Error(result.error.message));
            }
        });
    });
}

async function saveToBlob(emailData) {
    try {
        const response = await fetch(`${config.api.baseUrl}${config.api.endpoints.saveEmail}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(emailData)
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.detail || 'Failed to save email');
        }

        const result = await response.json();
        return result;
    } catch (error) {
        throw new Error('Error saving email: ' + error.message);
    }
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = type;
}
