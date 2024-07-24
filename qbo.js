const puppeteer = require("puppeteer-extra");
const { app } = require("@azure/functions");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const { BlobServiceClient } = require('@azure/storage-blob');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');

puppeteer.use(StealthPlugin());

app.http('qbo', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    context.log("Function 'qbo' started.");

    async function captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step) {
      try {
        const timeStamp = Date.now();
        const screenshotBuffer = await page.screenshot({ fullPage: true });
        const htmlContent = await page.content();

        const screenshotBlobName = `screenshots/${sanitizedWebsite}_${timeStamp}_${step}.png`;
        const screenshotBlockBlobClient = containerClient.getBlockBlobClient(screenshotBlobName);
        await screenshotBlockBlobClient.upload(screenshotBuffer, screenshotBuffer.length);

        const htmlBlobName = `html/${sanitizedWebsite}_${timeStamp}_${step}.txt`;
        const htmlBlockBlobClient = containerClient.getBlockBlobClient(htmlBlobName);
        await htmlBlockBlobClient.upload(htmlContent, Buffer.byteLength(htmlContent));

        context.log(`Screenshot and HTML content uploaded to Azure Blob Storage at step: ${step}`);
      } catch (captureError) {
        context.log(`Error capturing or uploading state at step ${step}:`, captureError.message);
      }
    }

    async function deleteBlobsInFolder(containerClient, folder) {
      try {
        for await (const blob of containerClient.listBlobsFlat({ prefix: folder })) {
          await containerClient.deleteBlob(blob.name);
          context.log(`Deleted blob: ${blob.name}`);
        }
      } catch (error) {
        context.log(`Error deleting blobs in folder ${folder}: ${error.message}`);
      }
    }

    async function sendEmail(to, cc, subject, text) {
      let transporter = nodemailer.createTransport({
        host: 'smtp.office365.com',
        port: 587,
        secure: false,
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

    async function monitorBlobStorage(containerClient, directory, timeout) {
      const endTime = Date.now() + timeout;
      while (Date.now() < endTime) {
        context.log("Checking blob storage for verification code...");
        let blobs = [];
        for await (const blob of containerClient.listBlobsFlat({ prefix: directory })) {
          blobs.push(blob.name);
        }

        blobs = blobs.filter(name => name !== `${directory}/readme.txt`);

        if (blobs.length > 0) {
          context.log("Found verification code file:", blobs[0]);
          return blobs[0].match(/\d{6}/)[0];
        }

        await new Promise(resolve => setTimeout(resolve, 10000));
      }
      throw new Error("Verification code not found within the timeout period");
    }

    function delay(time) {
      return new Promise(resolve => setTimeout(resolve, time));
    }

    let page;
    let step = 'init';

    try {
      context.log("Starting qbo function");
      context.log(request);

      const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING);
      const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_CONTAINER_NAME);

      // Delete existing blobs in /screenshots and /html folders
      await deleteBlobsInFolder(containerClient, 'screenshots/');
      await deleteBlobsInFolder(containerClient, 'html/');

      // Extract URL from query parameters or request body
      const url = `qbo.intuit.com/app/yourAccount?tab=billingdetails`;
      context.log(`Using URL: ${url}`);

      const browser = await puppeteer.launch({
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized'],
        headless: true,
        defaultViewport: null // This allows Puppeteer to use the maximum screen size
      });

      const utcNow = Date.now();
      const sanitizedWebsite = url.replace(/[^a-zA-Z0-9]/g, '_');

      page = await browser.newPage();
      await page.setViewport({ width: 1920, height: 1080 }); // Set viewport to a larger size

      // Set download behavior
      const downloadPath = path.resolve('/mnt/data/downloads');
      await page._client().send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath });

      context.log("Navigating to URL...");
      step = 'goto_url';
      await page.goto(`https://${url}`);
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

      context.log("Waiting for page to load...");
      step = 'wait_page_load';
      await delay(5000); // Wait for 5 seconds
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

      // Check if the page has loaded properly by looking for a specific element
      try {
        const userFieldSelector = '#iux-identifier-first-international-email-user-id-input';
        await page.waitForSelector(userFieldSelector, { timeout: 5000 });
        context.log("Page loaded. Filling in the user field.");
        step = 'fill_user_field';
        await page.type(userFieldSelector, process.env.QBO_USER);
        await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

        context.log("User field filled. Clicking the submit button.");
        step = 'click_submit_button';
        const submitButtonSelector = '[data-testid="IdentifierFirstSubmitButton"]';
        await page.waitForSelector(submitButtonSelector);
        await page.click(submitButtonSelector);
        await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

        context.log("Waiting for password field...");
        step = 'wait_password_field';
        const passwordFieldSelector = '#iux-password-confirmation-password';
        await page.waitForSelector(passwordFieldSelector, { timeout: 60000 });
        await page.type(passwordFieldSelector, process.env.QBO_PASSWORD);
        await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

        context.log("Clicking the continue button.");
        step = 'click_continue_button';
        const continueButtonSelector = '[data-testid="passwordVerificationContinueButton"]';
        await page.waitForSelector(continueButtonSelector);
        await page.click(continueButtonSelector);
        await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

        context.log("Waiting for page to load after button click...");
        step = 'wait_after_continue';
        await delay(5000); // Wait for 5 seconds
        await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

        context.log("Clicking text code button.");
        step = 'click_text_code_button';
        const textCodeButtonSelector = '[data-testid="challengePickerOption_SMS_OTP"]';
        await page.waitForSelector(textCodeButtonSelector, { timeout: 5000 });

        await page.click(textCodeButtonSelector);
        await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

        context.log("Sending verification email.");
        step = 'send_verification_email';
        await sendEmail(
          process.env.QBO_USER,
          process.env.QBO_CC_USER,
          'Your QBO Verification Code',
          'Please reply to this email with your verification code (in the email body - six digit code only, no text). If it has been more than 5 minutes since the email was sent please email DataAnalytics@bmss.com letting us know to rerun the process.'
        );
        await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

        let verificationCode;
        try {
          context.log("Retrieving verification code from blob storage.");
          step = 'retrieve_verification_code';
          verificationCode = await monitorBlobStorage(containerClient, 'qboCode', 300000);
          context.log("Verification code received:", verificationCode);
          await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

          context.log("Deleting blobs in qboCode directory except readme.txt...");
          for await (const blob of containerClient.listBlobsFlat({ prefix: 'qboCode/' })) {
            if (blob.name !== 'qboCode/readme.txt') {
              await containerClient.deleteBlob(blob.name);
              context.log(`Deleted blob: ${blob.name}`);
            }
          }
        } catch (error) {
          context.log("Error during verification code retrieval:", error.message);
          await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);
          context.res = {
            status: 500,
            body: `An error occurred during verification code retrieval: ${error.message}`
          };
          return;
        }

        context.log("Typing verification code into the input field.");
        step = 'input_verification_code';
        const verificationCodeInputSelector = '#ius-mfa-confirm-code';
        await page.waitForSelector(verificationCodeInputSelector);
        await page.type(verificationCodeInputSelector, verificationCode);
        await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

        context.log("Clicking verify continue button.");
        step = 'click_verify_continue';
        const verifyContinueButtonSelector = '[data-testid="VerifyOtpSubmitButton"]';
        await page.waitForSelector(verifyContinueButtonSelector);
        await page.click(verifyContinueButtonSelector);
        await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

        context.log("Waiting for page to load after verification.");
        step = 'wait_after_verification';
        await delay(5000); // Wait for 5 seconds
        await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);
      } catch (error) {
        context.log(`Error during steps before 'click_bmss_llc': ${error.message}`);
        await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);
      }

      context.log("Clicking BMSS, LLC button.");
      step = 'click_bmss_llc';
      const bmssButtonSelector = 'button.account-btn.account-btn-focus-quickbooks';
      await page.waitForSelector(bmssButtonSelector);
      const buttons = await page.$$(bmssButtonSelector);
      for (const button of buttons) {
        const accountName = await button.$eval('.account-name', el => el.textContent.trim());
        if (accountName === 'BMSS, LLC') {
          await button.click();
          break;
        }
      }
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

      context.log("Waiting for page to load after clicking BMSS, LLC button.");
      step = 'wait_after_bmss';
      await delay(20000); // Wait for 20 seconds
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

      context.log("Clicking billing details tab.");
      step = 'click_billing_details';
      const billingDetailsTabSelector = 'button#idsTab-tab_billing_details';
      await page.waitForSelector(billingDetailsTabSelector, { visible: true, timeout: 20000 }); // Wait until the button is visible
      await page.click(billingDetailsTabSelector);
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

      context.log("Waiting for billing details page to load.");
      step = 'wait_billing_details';
      await page.waitForFunction(() => !document.body.innerText.includes('Loading...'), { timeout: 300000 });
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

      context.log("Ensuring the correct row is targeted.");
      step = 'ensure_correct_row';
      const tableSelector = 'table';
      await page.waitForSelector(tableSelector, { timeout: 10000 });
      await page.waitForFunction(() => !document.body.innerText.includes('Loading...'), { timeout: 300000 });
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

      const rows = await page.$$('tbody > tr:not(.idsTable__headerRow)');
      await page.waitForFunction(() => !document.body.innerText.includes('Loading...'), { timeout: 300000 });

      if (rows.length > 0) {
          const firstRow = rows[0];
          await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

          context.log("Clicking expand menu button in the targeted row.");
          step = 'click_expand_menu';
          const firstRowExpandMenuButtonSelector = '[data-testid="chevron-down-icon-control"]';
          await page.waitForSelector(firstRowExpandMenuButtonSelector, { timeout: 10000 });
          await page.waitForFunction(() => !document.body.innerText.includes('Loading...'), { timeout: 300000 });
          const expandButton = await firstRow.$(firstRowExpandMenuButtonSelector);
          if (expandButton) {
              await expandButton.click();
              await delay(20000);
              await page.waitForFunction(() => !document.body.innerText.includes('Loading...'), { timeout: 300000 });
              await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

              context.log("Clicking export to CSV option.");
              step = 'click_export_csv';
              const exportCsvOptionSelector = 'li.Menu-menu-item-container-4cf029d';
              await page.waitForSelector(exportCsvOptionSelector, { timeout: 20000 });
              const exportCsvOption = await page.$(exportCsvOptionSelector);
              if (exportCsvOption) {
                  await exportCsvOption.click();
                  await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

                  // Wait for the file to be downloaded
                  context.log("Waiting for file to download...");
                  await delay(30000);// Adjust the timeout as necessary

                  const files = fs.readdirSync(downloadPath);
                  const csvFile = files.find(file => file.endsWith('.csv'));
                  if (csvFile) {
                      const filePath = path.join(downloadPath, csvFile);
                      const csvBlobName = `qboBillings/${csvFile}`;
                      const csvBlockBlobClient = containerClient.getBlockBlobClient(csvBlobName);

                      await csvBlockBlobClient.uploadFile(filePath);
                      context.log(`CSV file uploaded to Azure Blob Storage as ${csvBlobName}`);

                      // Clean up downloaded file
                      fs.unlinkSync(filePath);
                  } else {
                      throw new Error("CSV file not found in download directory.");
                  }
              } else {
                throw new Error("Export to CSV option not found.");
                  }
          } else {
              throw new Error("Expand menu button not found in the first row.");
          }
      } else {
          throw new Error("No rows found in the table.");
      }

      context.log("Waiting for final page to load.");
      step = 'wait_final_page';
      await delay(20000); // Wait for 20 seconds
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step);

      const finalScreenshotBuffer = await page.screenshot();
      const finalHtmlContent = await page.content();

      const finalScreenshotBlobName = `screenshots/final_${sanitizedWebsite}_${utcNow}.png`;
      const finalScreenshotBlockBlobClient = containerClient.getBlockBlobClient(finalScreenshotBlobName);
      await finalScreenshotBlockBlobClient.upload(finalScreenshotBuffer, finalScreenshotBuffer.length);

      const finalHtmlBlobName = `html/final_html_${sanitizedWebsite}_${utcNow}.txt`;
      const finalHtmlBlockBlobClient = containerClient.getBlockBlobClient(finalHtmlBlobName);
      await finalHtmlBlockBlobClient.upload(finalHtmlContent, Buffer.byteLength(finalHtmlContent));

      context.log("Final screenshot and HTML content uploaded to Azure Blob Storage");

      context.res = {
        status: 200,
        body: finalHtmlContent,
        headers: {
          'Content-Type': 'text/html'
        }
      };
    } catch (error) {
      context.log("Error during function execution:", error.message);
      context.res = {
        status: 500,
        body: `An error occurred during processing at step ${step}: ${error.message}`
      };
    }
  }
});
