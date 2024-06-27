const puppeteer = require("puppeteer-extra");
const { app } = require("@azure/functions");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const { BlobServiceClient } = require('@azure/storage-blob');

puppeteer.use(StealthPlugin());

app.http('qbo', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    context.log("Function 'qbo' started.");

    try {
      context.log("Starting qbo function");
      context.log(request);

      // Extract URL from query parameters or request body
      const url = `qbo.intuit.com`;
      context.log(`Using URL: ${url}`);

      const browser = await puppeteer.launch({
        args: ['--no-sandbox', '--disable-setuid-sandbox'],
        headless: true
      });

      const utcNow = Date.now();
      const sanitizedWebsite = url.replace(/[^a-zA-Z0-9]/g, '_');

      const page = await browser.newPage();
      context.log("Navigating to URL...");
      await page.goto(`https://${url}`);

      // Add a wait period to ensure the page has fully loaded
      context.log("Waiting for page to load...");
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for 5 seconds

      context.log("Page loaded. Filling in the user field.");

      // Enter QBO_USER into the email field
      const userFieldSelector = '#iux-identifier-first-international-email-user-id-input';
      await page.waitForSelector(userFieldSelector);
      context.log(`Typing into user field with selector: ${userFieldSelector}`);
      await page.type(userFieldSelector, process.env.QBO_USER);

      context.log("User field filled. Clicking the submit button.");

      // Click the submit button
      const submitButtonSelector = '[data-testid="IdentifierFirstSubmitButton"]';
      await page.waitForSelector(submitButtonSelector);
      context.log(`Found submit button with selector: ${submitButtonSelector}`);
      await page.click(submitButtonSelector);
      context.log("Submit button clicked.");

      // Wait for the password field to appear
      const passwordFieldSelector = '#iux-password-confirmation-password';
      context.log("Waiting for password field...");
      await page.waitForSelector(passwordFieldSelector, { timeout: 60000 }); // 60 seconds timeout
      context.log(`Typing into password field with selector: ${passwordFieldSelector}`);
      await page.type(passwordFieldSelector, process.env.QBO_PASSWORD);

      const continueButtonSelector = '[data-testid="passwordVerificationContinueButton"]';
      await page.waitForSelector(continueButtonSelector);
      context.log(`Found continue button with selector: ${continueButtonSelector}`);
      await page.click(continueButtonSelector);
      context.log("Continue button clicked.");

      const textCodeButtonSelector = '[data-testid="challengePickerOption_SMS_OTP"]';
      await page.waitForSelector(textCodeButtonSelector);
      context.log(`Found text code button with selector: ${textCodeButtonSelector}`);
      await page.click(textCodeButtonSelector);
      context.log("Text code button clicked.");

      // Add a wait period to ensure the page has fully loaded
      context.log("Waiting for page to load after clicking text code button...");
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for 5 seconds

      // Capture the screenshot of the next page
      const screenshotBuffer = await page.screenshot();

      // Get the HTML content of the next page
      const htmlContent = await page.content();
      await browser.close();

      const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING);
      const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_CONTAINER_NAME);

      // Save the screenshot
      const screenshotBlobName = `screenshots/${sanitizedWebsite}_${utcNow}.png`;
      const screenshotBlockBlobClient = containerClient.getBlockBlobClient(screenshotBlobName);
      await screenshotBlockBlobClient.upload(screenshotBuffer, screenshotBuffer.length);

      // Save the HTML content as a .txt file
      const htmlBlobName = `html/html_${sanitizedWebsite}_${utcNow}.txt`;
      const htmlBlockBlobClient = containerClient.getBlockBlobClient(htmlBlobName);
      await htmlBlockBlobClient.upload(htmlContent, Buffer.byteLength(htmlContent));

      context.log("Screenshot and HTML content uploaded to Azure Blob Storage");

      // Return the HTML content in the response
      context.res = {
        status: 200,
        body: htmlContent,
        headers: {
          'Content-Type': 'text/html'
        }
      };
    } catch (error) {
      context.log("Error during function execution:", error.message);

      context.res = {
        status: 500,
        body: `An error occurred during processing: ${error.message}`
      };
    }
  }
});
