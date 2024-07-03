const puppeteer = require("puppeteer-extra");
const { app } = require("@azure/functions");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const { BlobServiceClient } = require('@azure/storage-blob');
const fs = require('fs');
const path = require('path');

puppeteer.use(StealthPlugin());

async function uploadScreenshot(page, blobServiceClient, containerClient, sanitizedWebsite, step, utcNow, context) {
  const screenshotBuffer = await page.screenshot();
  const screenshotBlobName = `screenshots/${sanitizedWebsite}_${utcNow}_step${step}.png`;
  const screenshotBlockBlobClient = containerClient.getBlockBlobClient(screenshotBlobName);
  await screenshotBlockBlobClient.upload(screenshotBuffer, screenshotBuffer.length);
  context.log(`Screenshot for step ${step} uploaded to Azure Blob Storage`);
  return screenshotBuffer;
}

async function uploadHtmlContent(page, blobServiceClient, containerClient, sanitizedWebsite, step, utcNow, context) {
  const htmlContent = await page.content();
  const htmlBlobName = `html/html_${sanitizedWebsite}_${utcNow}_step${step}.html`;
  const htmlBlockBlobClient = containerClient.getBlockBlobClient(htmlBlobName);
  await htmlBlockBlobClient.upload(htmlContent, Buffer.byteLength(htmlContent));
  context.log(`HTML content for step ${step} uploaded to Azure Blob Storage`);
  return htmlContent;
}

async function uploadFileToBlob(filePath, blobServiceClient, containerClient, fileBlobName, context) {
  const fileBuffer = fs.readFileSync(filePath);
  const fileBlockBlobClient = containerClient.getBlockBlobClient(fileBlobName);
  await fileBlockBlobClient.upload(fileBuffer, fileBuffer.length);
  context.log("Downloaded file uploaded to Azure Blob Storage");
}

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

    const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING);
    const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_CONTAINER_NAME);

    const browser = await puppeteer.launch({
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
      headless: true
    });

    const utcNow = Date.now();
    const sanitizedWebsite = url.replace(/[^a-zA-Z0-9]/g, '_');
    const downloadPath = path.resolve('/tmp', `download_${utcNow}`);
    fs.mkdirSync(downloadPath, { recursive: true });

    let page, screenshotBuffer, htmlContent;

    try {
      page = await browser.newPage();

      // Set the download behavior to allow downloads to a specified directory
      await page._client().send('Page.setDownloadBehavior', {
        behavior: 'allow',
        downloadPath: downloadPath
      });
      context.log("Download behavior set");

      await page.goto(url);
      context.log("Navigated to login page");

      await page.type('#UserName', username);
      await page.type('#Password', password);
      await page.click('#LoginButton');
      context.log("Login credentials entered and login button clicked");

      // Add a wait period to ensure the page has fully loaded after login
      await new Promise(resolve => setTimeout(resolve, 10000)); // Wait for 10 seconds
      context.log("Waited for 10 seconds after login");

      // Capture screenshot and HTML after login
      screenshotBuffer = await uploadScreenshot(page, blobServiceClient, containerClient, sanitizedWebsite, 'login', utcNow, context);
      htmlContent = await uploadHtmlContent(page, blobServiceClient, containerClient, sanitizedWebsite, 'login', utcNow, context);

      // Click on the "Reports" tab
      await page.click('#ahrefreports_menu');
      context.log("Clicked on 'Reports' tab");

      // Add a wait period to ensure the reports page has fully loaded
      await new Promise(resolve => setTimeout(resolve, 10000)); // Wait for 10 seconds
      context.log("Waited for 10 seconds after clicking 'Reports' tab");

      // Capture screenshot and HTML after clicking Reports tab
      screenshotBuffer = await uploadScreenshot(page, blobServiceClient, containerClient, sanitizedWebsite, 'reports_tab', utcNow, context);
      htmlContent = await uploadHtmlContent(page, blobServiceClient, containerClient, sanitizedWebsite, 'reports_tab', utcNow, context);

      // Switch to the frame containing the actual report content
      const frame = page.frames().find(frame => frame.name() === 'main');
      if (!frame) throw new Error('Frame "main" not found');

      context.log("Switched to frame 'main'");

      // Wait for the Export button to be visible and clickable
      await frame.waitForSelector('#bexport', { visible: true });
      context.log("Export button is visible");

      // Click on the "Export" button
      await frame.click('#bexport');
      context.log("Clicked on 'Export' button");

      // Add a wait period to ensure the download has started and completed
      await new Promise(resolve => setTimeout(resolve, 10000)); // Wait for 10 seconds
      context.log("Waited for 10 seconds after clicking 'Export' button");

      // Capture screenshot and HTML after clicking Export button
      //screenshotBuffer = await uploadScreenshot(frame, blobServiceClient, containerClient, sanitizedWebsite, 'export', utcNow, context);
      //htmlContent = await uploadHtmlContent(frame, blobServiceClient, containerClient, sanitizedWebsite, 'export', utcNow, context);

      // Get the HTML content of the page
      htmlContent = await frame.content();
      context.log("HTML content captured");

      // Find the downloaded file
      const files = fs.readdirSync(downloadPath);
      const downloadedFile = files.find(file => file.endsWith('.xlsx'));
      const filePath = path.join(downloadPath, downloadedFile);
      context.log("Downloaded file found:", downloadedFile);

      // Save the downloaded file to Azure Blob Storage
      const fileBlobName = `inbound/eSuite/employeeProject_${utcNow}.xlsx`;
      await uploadFileToBlob(filePath, blobServiceClient, containerClient, fileBlobName, context);

    } catch (error) {
      context.log("Error during function execution:", error.message);
      htmlContent = "An error occurred during processing.";
      screenshotBuffer = Buffer.from([]);
    } finally {
      await browser.close();
      context.log("Browser closed");
    }

    //try {
      // Save the final HTML content as a .txt file
    //  await uploadHtmlContent(page, blobServiceClient, containerClient, sanitizedWebsite, 'final', utcNow, context);
    //} catch (error) {
    //  context.log("Error during blob storage upload:", error.message);
    //}

    context.res = {
      status: 200,
      body: htmlContent,
      headers: {
        'Content-Type': 'text/html'
      }
    };
  }
});
