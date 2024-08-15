const puppeteer = require("puppeteer-extra");
const { app } = require("@azure/functions");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const { BlobServiceClient } = require('@azure/storage-blob');

puppeteer.use(StealthPlugin());

async function captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step, context) {
  try {
    const timeStamp = Date.now();
    const screenshotBuffer = await page.screenshot({ fullPage: true });
    const htmlContent = await page.content();

    const screenshotBlobName = `screenshots/${sanitizedWebsite}_${timeStamp}_${step}.png`;
    const screenshotBlockBlobClient = containerClient.getBlockBlobClient(screenshotBlobName);
    await screenshotBlockBlobClient.upload(screenshotBuffer, screenshotBuffer.length);

    const htmlBlobName = `html/${sanitizedWebsite}_${timeStamp}_${step}.html`;
    const htmlBlockBlobClient = containerClient.getBlockBlobClient(htmlBlobName);
    await htmlBlockBlobClient.upload(htmlContent, Buffer.byteLength(htmlContent));

    context.log(`Screenshot and HTML content uploaded to Azure Blob Storage at step: ${step}`);
  } catch (captureError) {
    context.log(`Error capturing or uploading state at step ${step}:`, captureError.message);
  }
}

app.http('lcVista', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    context.log("Function 'lcVista' started.");

    const url = "https://bmssu.lcvista.com/";
    const username = process.env.LCVISTA_USERNAME; // Assuming the username is stored in an environment variable
    const password = process.env.LCVISTA_PASSWORD; // Assuming the password is stored in an environment variable

    const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING);
    const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_CONTAINER_NAME);

    const browser = await puppeteer.launch({
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
      headless: true
    });

    const utcNow = Date.now();
    const sanitizedWebsite = url.replace(/[^a-zA-Z0-9]/g, '_');
    function delay(time) {
        return new Promise(resolve => setTimeout(resolve, time));
      };

    let step = 'initial';

    try {
      const page = await browser.newPage();
      await page.goto(url, { waitUntil: 'networkidle0' });
      context.log(`Navigated to ${url}`);

      // Capture screenshot and HTML content after the initial load
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step, context);

      // Step 2: Fill in the username field
      step = 'fill_username';
      await page.type('#i0116', username);
      context.log("Filled in the username field");

      // Capture screenshot and HTML content after filling in the username
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step, context);

      // Step 3: Click the submit button (after username)
      step = 'click_submit_button_username';
      await page.click('#idSIButton9');
      context.log("Clicked the submit button after username");
      await delay(5000); // Wait for 5 seconds


      // Capture screenshot and HTML content after clicking the submit button
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step, context);

      // Step 4: Fill in the password field
      step = 'fill_password';
      await page.waitForSelector('#i0118');  // Ensure the password field is available
      await page.type('#i0118', password);
      context.log("Filled in the password field");
      await delay(5000); // Wait for 5 seconds


      // Capture screenshot and HTML content after filling in the password
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step, context);

      // Step 5: Click the submit button (after password)
      step = 'click_submit_button_password';
      await page.click('#idSIButton9');
      context.log("Clicked the submit button after password");
      await delay(5000); // Wait for 5 seconds


      // Capture screenshot and HTML content after clicking the submit button
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step, context);

      // Step 6: Click the "No" button
      step = 'click_no_button';
      await page.waitForSelector('#idBtn_Back');  // Ensure the "No" button is available
      await page.click('#idBtn_Back');
      context.log("Clicked the 'No' button");
      await delay(5000); // Wait for 5 seconds


      // Capture screenshot and HTML content after clicking the "No" button
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step, context);

      // Step 7: Click the "Reports" link
      step = 'click_reports_link';
      await page.waitForSelector('a[href="/bmssu/reports/"]');  // Ensure the "Reports" link is available
      await page.click('a[href="/bmssu/reports/"]');
      context.log("Clicked the 'Reports' link");
      await delay(5000); // Wait for 5 seconds


      // Capture screenshot and HTML content after clicking the "Reports" link
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step, context);

    } catch (error) {
      context.log("Error during function execution:", error.message);
    } finally {
      await browser.close();
      context.log("Browser closed");
    }

    context.res = {
      status: 200,
      body: "All steps completed, and content uploaded.",
      headers: {
        'Content-Type': 'text/plain'
      }
    };
  }
});
