# Outlook Training Email Saver

This Outlook add-in adds a "Send and Save for Training" button that sends emails and automatically saves them to Azure Blob Storage for training purposes.

## Setup Instructions

1. Install dependencies:
```bash
npm install
```

2. Configure Azure Storage:
   - Create an Azure Storage account and container
   - Update the connection string in `src/config.js`
   - Create an Azure Function to handle the blob storage upload (for security)

3. Generate development certificates:
```bash
npm run office-addin-dev-certs install
```

4. Start the development server:
```bash
npm start
```

5. Sideload the add-in in Outlook:
   - Go to Outlook Settings
   - Select "Get Add-ins"
   - Choose "My Add-ins"
   - Select "Add Custom Add-in" and choose the manifest.xml file

## Features

- Adds a "Send and Save for Training" button to the email compose window
- Automatically sends the email and saves it to Azure Blob Storage
- Saves email metadata including subject, body, sender, recipients, and timestamp
- Provides visual feedback on success/failure

## Security Notes

- The add-in uses Azure Functions to securely handle blob storage operations
- Never expose storage connection strings directly in client-side code
- Implement proper authentication in your Azure Function

## Development

To start development server with file watching:
```bash
npm run watch
```

To validate the manifest:
```bash
npm run validate
```
