const puppeteer = require("puppeteer-extra");
const { app } = require("@azure/functions");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const { BlobServiceClient } = require('@azure/storage-blob');

puppeteer.use(StealthPlugin());

app.http('eSuite', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    context.log("Function 'eSuite' started.");

    const url = "https://www.empiresuite.com/login?sso=skip";
    const username = process.env.EMPIRESUITE_USER;
    const password = process.env.EMPIRESUITE_PW;

    if (!username || !password) {
      context.res = {
        status: 400,
        body: "Username or password environment variable is missing"
      };
      return;
    }

    const browser = await puppeteer.launch({
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
      headless: true
    });

    const utcNow = Date.now();
    const sanitizedWebsite = url.replace(/[^a-zA-Z0-9]/g, '_');
    let page, screenshotBuffer, htmlContent;

    try {
      page = await browser.newPage();
      await page.goto(url);

      await page.type('#UserName', username);
      await page.type('#Password', password);
      await page.click('#LoginButton');

      // Add a wait period to ensure the page has fully loaded
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for 5 seconds

      // Capture the screenshot
      screenshotBuffer = await page.screenshot();

      // Get the HTML content of the page
      htmlContent = await page.content();
      
    } catch (error) {
      context.log("Error during function execution:", error.message);
      htmlContent = "An error occurred during processing.";
      screenshotBuffer = Buffer.from([]);
    } finally {
      await browser.close();
    }

    try {
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

    } catch (error) {
      context.log("Error during blob storage upload:", error.message);
    }

    context.res = {
      status: 200,
      body: htmlContent,
      headers: {
        'Content-Type': 'text/html'
      }
    };
  }
});
