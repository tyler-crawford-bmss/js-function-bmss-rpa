const puppeteer = require("puppeteer-extra");
const { app } = require("@azure/functions");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const { BlobServiceClient } = require('@azure/storage-blob');
const nodemailer = require('nodemailer');

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

      // Add a wait period after the button click
      context.log("Waiting for the page to load after button click...");
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for 5 seconds

      const textCodeButtonSelector = '[data-testid="challengePickerOption_SMS_OTP"]';
      await page.waitForSelector(textCodeButtonSelector);
      context.log(`Found text code button with selector: ${textCodeButtonSelector}`);
      await page.click(textCodeButtonSelector);
      context.log("Text code button clicked.");

      // Send email to QBO_USER
      async function sendEmail(to, cc, subject, text) {
        let transporter = nodemailer.createTransport({
          host: 'smtp.office365.com',
          port: 587,
          secure: false, // true for 465, false for other ports
          auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASS
          }
        });

        let info = await transporter.sendMail({
          from: process.env.EMAIL_USER,
          to: to,
          cc: cc,
          subject: subject,
          html: `${text}<br><br><img src="https://storageacctbmssprod001.blob.core.windows.net/container-bmssprod001-public/images/KB-Signature.png" alt="Signature" />`
        });

        context.log('Message sent: %s', info.messageId);
      }

      // Call this function after clicking the text code button
      await sendEmail(
        process.env.QBO_USER,
        process.env.QBO_CC_USER,
        'Your QBO Verification Code',
        'Please reply to this email with your verification code (in the email body - six digit code only, no text). If it has been more than 5 minutes since the email was sent please email DataAnalytics@bmss.com letting us know to rerun the process.'
      );
      context.log("Verification email sent to QBO_USER.");

      // Function to monitor Azure Blob Storage for the verification code file
      async function monitorBlobStorage(containerClient, directory, timeout) {
        const endTime = Date.now() + timeout;
        while (Date.now() < endTime) {
          context.log("Checking blob storage for verification code...");
          let blobs = [];
          for await (const blob of containerClient.listBlobsFlat({ prefix: directory })) {
            blobs.push(blob.name);
          }

          // Exclude the readme.txt file
          blobs = blobs.filter(name => name !== `${directory}/readme.txt`);

          if (blobs.length > 0) {
            context.log("Found verification code file:", blobs[0]);
            return blobs[0].match(/\d{6}/)[0]; // Extract the 6-digit code from the file name
          }

          // Wait for 10 seconds before checking again
          await new Promise(resolve => setTimeout(resolve, 10000));
        }
        throw new Error("Verification code not found within the timeout period");
      }

      const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING);
      const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_CONTAINER_NAME);

      let verificationCode;
      try {
        verificationCode = await monitorBlobStorage(containerClient, 'qboCode', 300000); // 5 minutes timeout
        context.log("Verification code received:", verificationCode);

        // Delete all blobs in the qboCode directory except readme.txt
        context.log("Deleting blobs in qboCode directory except readme.txt...");
        for await (const blob of containerClient.listBlobsFlat({ prefix: 'qboCode/' })) {
          if (blob.name !== 'qboCode/readme.txt') {
            await containerClient.deleteBlob(blob.name);
            context.log(`Deleted blob: ${blob.name}`);
          }
        }
      } catch (error) {
        context.log("Error during verification code retrieval:", error.message);
        context.res = {
          status: 500,
          body: `An error occurred during verification code retrieval: ${error.message}`
        };
        return;
      }

      // Input the verification code into the form field
      const verificationCodeInputSelector = '#ius-mfa-confirm-code';
      await page.waitForSelector(verificationCodeInputSelector);
      context.log(`Typing verification code into the input field with selector: ${verificationCodeInputSelector}`);
      await page.type(verificationCodeInputSelector, verificationCode);

      // Click the "Continue" button
      const verifyContinueButtonSelector = '[data-testid="VerifyOtpSubmitButton"]';
      await page.waitForSelector(verifyContinueButtonSelector);
      context.log(`Found continue button with selector: ${verifyContinueButtonSelector}`);
      await page.click(verifyContinueButtonSelector);
      context.log("Verification code submitted and Continue button clicked.");

      // Wait for the page to load after clicking the Continue button
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for 5 seconds

      // Select the "BMSS, LLC" button
      const bmssButtonSelector = 'button.account-btn.account-btn-focus-quickbooks';
      await page.waitForSelector(bmssButtonSelector);
      const buttons = await page.$$(bmssButtonSelector);
      for (const button of buttons) {
        const accountName = await button.$eval('.account-name', el => el.textContent.trim());
        if (accountName === 'BMSS, LLC') {
          context.log('Clicking BMSS, LLC button');
          await button.click();
          break;
        }
      }
      context.log("BMSS, LLC button clicked.");

      // Wait for the final page to load
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for 5 seconds

      // Capture the screenshot of the final page
      const screenshotBuffer = await page.screenshot();

      // Get the HTML content of the final page
      const htmlContent = await page.content();
      await browser.close();

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
