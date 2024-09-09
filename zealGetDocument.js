const puppeteer = require("puppeteer-extra");
const { app } = require("@azure/functions");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const { BlobServiceClient } = require('@azure/storage-blob');
const fs = require('fs');
const path = require('path');
const mime = require('mime'); // to detect file type and extension

puppeteer.use(StealthPlugin());

app.http('zealGetDocument', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
      context.log("Function 'zealGetDocument' started.");
  
      // Parse the JSON body using the `json()` method
      let reqBody;
      try {
        reqBody = await request.json(); // Parse the JSON body
      } catch (error) {
        context.log("Error parsing request body: ", error.message);
        context.res = {
          status: 400,
          body: 'Invalid JSON in request body.'
        };
        return;
      }
  
      // Log the request body
      context.log(`Request Body Looks like this: ${JSON.stringify(reqBody)}`);
  
      // Extract the path and documentName from the request body
      const { path: urlPath, documentName } = reqBody;
      context.log(`Path looks like this: ${urlPath}`);
      context.log(`Document name to save as: ${documentName}`);

      // Validate if path and documentName are present
      if (!urlPath || !documentName) {
        context.log("No path or documentName provided in the request body.");
        context.res = {
          status: 400,
          body: 'Error: No path or documentName provided in the request body.'
        };
        return;
      }

      async function savePageAsFile(page, downloadPath, documentName) {
        try {
          // Get the page content as a buffer (you can also download a specific link by modifying this)
          const response = await page.goto(urlPath, { waitUntil: 'networkidle2' });
          const fileBuffer = await response.buffer();

          // Determine the file type based on the content
          const mimeType = response.headers()['content-type'];
          const extension = mime.getExtension(mimeType) || 'pdf'; // Use 'pdf' as default if unknown
          const fileName = `${documentName}.${extension}`;
          const filePath = path.join(downloadPath, fileName);

          // Save the file locally
          fs.writeFileSync(filePath, fileBuffer);
          context.log(`File saved as: ${filePath}`);

          return filePath;
        } catch (error) {
          context.log(`Error saving the document: ${error.message}`);
          throw error;
        }
      }

      async function uploadDocumentToBlob(containerClient, filePath, blobName) {
        try {
          const blockBlobClient = containerClient.getBlockBlobClient(blobName);
          await blockBlobClient.uploadFile(filePath);
          context.log(`File uploaded to Azure Blob Storage as ${blobName}`);
        } catch (error) {
          context.log(`Error uploading file to Azure Blob Storage: ${error.message}`);
          throw error;
        }
      }

      let page;
      let step = 'init';

      try {
        const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING);
        const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_CONTAINER_NAME);

        context.log("Launching the browser...");
        const browser = await puppeteer.launch({
          args: ['--no-sandbox', '--disable-setuid-sandbox'],
          headless: true, // Set to true if you want headless mode
          defaultViewport: null
        });

        const utcNow = Date.now();
        const sanitizedWebsite = 'zeal_bmss_com';

        context.log("Opening a new browser page...");
        page = await browser.newPage();
        await page.setViewport({ width: 1920, height: 1080 });

        context.log("Navigating to Zeal document page...");
        step = 'goto_zeal';
        await page.goto(urlPath);
        context.log("Page navigation completed.");

        context.log("Waiting for 30 seconds...");
        await new Promise(resolve => setTimeout(resolve, 30000)); // Wait for 30 seconds

        // Define the download path (temporary storage in Azure Functions)
        const downloadPath = path.resolve('/mnt/data/downloads');
        if (!fs.existsSync(downloadPath)) {
          fs.mkdirSync(downloadPath, { recursive: true });
        }

        // Simulate a "Save As" by downloading the page as a file
        const downloadedFilePath = await savePageAsFile(page, downloadPath, documentName);

        // Upload the downloaded document to Azure Blob Storage
        const blobName = `screenshots/${documentName}${path.extname(downloadedFilePath)}`;
        await uploadDocumentToBlob(containerClient, downloadedFilePath, blobName);

        context.log("Closing the browser...");
        await browser.close();

        context.res = {
          status: 200,
          body: `Document downloaded and uploaded as ${blobName} successfully.`,
        };

      } catch (error) {
        context.log(`Error during function execution: ${error.message}`);
        context.res = {
          status: 500,
          body: `An error occurred during processing at step ${step}: ${error.message}`,
        };
      }
    }
});
