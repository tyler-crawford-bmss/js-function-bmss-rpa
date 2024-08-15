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

    const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING);
    const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_CONTAINER_NAME);

    const browser = await puppeteer.launch({
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
      headless: true
    });

    const utcNow = Date.now();
    const sanitizedWebsite = url.replace(/[^a-zA-Z0-9]/g, '_');
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

    } catch (error) {
      context.log("Error during function execution:", error.message);
    } finally {
      await browser.close();
      context.log("Browser closed");
    }

    context.res = {
      status: 200,
      body: "Username filled in and content uploaded.",
      headers: {
        'Content-Type': 'text/plain'
      }
    };
  }
});
