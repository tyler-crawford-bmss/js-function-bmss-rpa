const puppeteer = require("puppeteer-extra");
const { app } = require("@azure/functions");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const { BlobServiceClient } = require('@azure/storage-blob');
const xlsx = require('xlsx'); // Added to handle Excel files
const fs = require('fs');
const path = require('path');

puppeteer.use(StealthPlugin());

// Existing function to capture screenshots and HTML content
async function captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step, context) {
  try {
    const timeStamp = Date.now();
    const screenshotBuffer = await page.screenshot({ fullPage: true });
    
    // Get the main page content (frameset)
    let htmlContent = await page.content();

    // Iterate over all frames and replace frame src with actual content
    const frames = page.frames();
    for (const frame of frames) {
      if (frame.parentFrame() !== null) { // Exclude the main frame
        const frameName = frame.name();
        const frameContent = await frame.content();

        // Replace the frame's src attribute in the main HTML with the actual content
        const regex = new RegExp(`<frame[^>]*name="${frameName}"[^>]*src="[^"]*"[^>]*>`, 'i');
        const replacement = `<iframe name="${frameName}" srcdoc="${frameContent.replace(/"/g, '&quot;')}"></iframe>`;
        htmlContent = htmlContent.replace(regex, replacement);
      }
    }

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

// Function to copy blob from one path to another
async function copyBlob(blobServiceClient, sourceContainer, sourceBlob, destinationContainer, destinationBlob, context) {
    const sourceBlobClient = blobServiceClient.getContainerClient(sourceContainer).getBlobClient(sourceBlob);
    const destinationBlobClient = blobServiceClient.getContainerClient(destinationContainer).getBlobClient(destinationBlob);
    await destinationBlobClient.beginCopyFromURL(sourceBlobClient.url);
    context.log(`Copied blob from ${sourceBlob} to ${destinationBlob}`);
  }

// Function to upload files to Azure Blob Storage
async function uploadFileToBlob(filePath, blobServiceClient, containerClient, fileBlobName, context) {
  const fileBuffer = fs.readFileSync(filePath);
  const fileBlockBlobClient = containerClient.getBlockBlobClient(fileBlobName);
  await fileBlockBlobClient.upload(fileBuffer, fileBuffer.length);
  context.log(`File ${filePath} uploaded to Azure Blob Storage as ${fileBlobName}`);
}

// Function to convert Excel files to CSV
async function convertXlsxToCsv(xlsxFilePath, csvDirPath) {
  const workbook = xlsx.readFile(xlsxFilePath);
  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    const csv = xlsx.utils.sheet_to_csv(worksheet, { FS: '|' });
    const csvFilePath = path.join(csvDirPath, `${sheetName}.csv`);
    fs.writeFileSync(csvFilePath, csv);
  });
}

app.http('eSuite', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    context.log("Function 'eSuiteNew' started.");

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
      headless: true,
      slowMo: 50
    });

    const utcNow = Date.now();
    const sanitizedWebsite = url.replace(/[^a-zA-Z0-9]/g, '_');
    const downloadPath = path.resolve('/tmp', `download_${utcNow}`);
    fs.mkdirSync(downloadPath, { recursive: true });

    try {
      // Step 1: Copy existing CSV file to a new location
      const existingCsvFileName = 'inbound/eSuite/employeeProject.csv';
      const lastWeekCsvFileName = 'inbound/eSuite/employeeProject_lastWeek.csv';
      await copyBlob(blobServiceClient, process.env.AZURE_CONTAINER_NAME, existingCsvFileName, process.env.AZURE_CONTAINER_NAME, lastWeekCsvFileName, context);
      
      context.log("Existing CSV file saved as last week's file");  
      let page;

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

      // Wait for the page to load after login
      await new Promise(resolve => setTimeout(resolve, 30000)); // Wait for 30 seconds
      context.log("Waited for 30 seconds after login");

      // Capture screenshot and HTML after login
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, 'login', context);

      // Click on the "Reports" tab
      await page.click('#ahrefreports_menu');
      context.log("Clicked on 'Reports' tab");

      // Wait for the reports page to load
      await new Promise(resolve => setTimeout(resolve, 30000)); // Wait for 30 seconds
      context.log("Waited for 30 seconds after clicking 'Reports' tab");

      // Switch to the frame containing the actual report content
      let frame = page.frames().find(frame => frame.name() === 'main');
      if (!frame) throw new Error('Frame "main" not found');

      context.log("Switched to frame 'main'");

      // Set up a listener for the new page before clicking
      const newPagePromise = new Promise(resolve => {
        page.once('popup', popup => resolve(popup));
      });

      // Click on the "binoculars.gif" button (employee search)
      await frame.click('a[href="javascript:OnEmployeeSearch()"] img[src="binoculars.gif"]');
      context.log('Clicked on the "binoculars.gif" button');

      // Wait for the new page to be created
      const newPage = await newPagePromise;

      context.log('New page opened');

      // Wait for the new page to load
      await new Promise(resolve => setTimeout(resolve, 30000));
      context.log('Waited for 30 seconds after opening employee search page');

      // Capture screenshot and HTML of the new page
      await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'employee_search', context);

      // Get the "results" frame
      const resultsFrame = newPage.frames().find(frame => frame.name() === 'results');
      if (!resultsFrame) throw new Error('Frame "results" not found');
      context.log('Accessed the "results" frame');

      // Step 1: Check the header checkbox within the "results" frame
      await resultsFrame.click('tr.list-h input[type="checkbox"][onclick="return onSelectAll(this)"]');
      context.log('Step 1: Checked the header checkbox');
      await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step1', context);
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for 5 seconds

      // Step 2: Click the "Remove" button within the "results" frame
      await resultsFrame.click('input[type="button"][value="Remove"][onclick="onRemoveSelected(\'ShowSelectedList\')"]');
      context.log('Step 2: Clicked the "Remove" button');
      await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step2', context);
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for 5 seconds

      // For the dropdown in the "criteria" frame
      const criteriaFrame = newPage.frames().find(frame => frame.name() === 'criteria');
      if (!criteriaFrame) throw new Error('Frame "criteria" not found');
      context.log('Accessed the "criteria" frame');

      // Step 3: Select "Active" from the dropdown within the "criteria" frame
      await criteriaFrame.select('select[name="emplstatus"]', '1'); // Select the option with value="1"
      context.log('Step 3: Selected "Active" from the dropdown');
      await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step3', context);
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for 5 seconds

      // Step 4: Click the "Search" button within the "criteria" frame
      await criteriaFrame.click('input[type="button"][value="Search"]');
      context.log('Step 4: Clicked the "Search" button');
      await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step4', context);
      await new Promise(resolve => setTimeout(resolve, 10000)); // Wait for 10 seconds

      // Back to the "results" frame for the remaining steps
      // Step 5: Check the header checkbox again within the "results" frame
      await resultsFrame.click('tr.list-h input[type="checkbox"][onclick="return onSelectAll(this)"]');
      context.log('Step 5: Checked the header checkbox again');
      await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step5', context);
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for 5 seconds

      // Step 6: Click the "Add" button within the "results" frame
      await resultsFrame.click('input[type="button"][value="Add"][onclick="onAddSelected()"]');
      context.log('Step 6: Clicked the "Add" button');
      await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step6', context);
      await new Promise(resolve => setTimeout(resolve, 10000)); // Wait for 10 seconds

      // After clicking "Add", re-fetch the frames
      const framesAfterAdd = newPage.frames();
      framesAfterAdd.forEach(frame => {
        context.log(`Frame name: ${frame.name()}, URL: ${frame.url()}`);
      });

      // Try to find the frame that contains the "OK" button
      let okButtonFrame = null;

      for (const frame of framesAfterAdd) {
        const okButtonExists = await frame.$('input[type="button"][id="doIt"][value="OK"]');
        if (okButtonExists) {
          okButtonFrame = frame;
          context.log(`Found "OK" button in frame: ${frame.name()}`);
          break;
        }
      }

      if (okButtonFrame) {
        // Capture state before clicking "OK"
        await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step7_before_click', context);

        // Step 7: Click the "OK" button within the identified frame
        await okButtonFrame.click('input[type="button"][id="doIt"][value="OK"]');
        context.log('Step 7: Clicked the "OK" button');
        await new Promise(resolve => setTimeout(resolve, 10000)); // Wait for 10 seconds

        // Do not capture state here, as the page may have closed
      } else {
        throw new Error('Could not find the "OK" button after clicking "Add"');
      }

      // Now, back to the main page
      // The newPage might have closed automatically; ensure it's closed
      try {
        await newPage.close();
        context.log('Closed the employee search window');
      } catch (e) {
        context.log('The employee search window was already closed');
      }

      // Bring the main page to the front
      await page.bringToFront();
      context.log('Brought main page to the front');

      // Re-fetch the 'main' frame
      frame = page.frames().find(frame => frame.name() === 'main');
      if (!frame) throw new Error('Frame "main" not found after closing employee search window');

      context.log("Switched back to frame 'main'");

      // Wait for the Export button to be visible and clickable
      await new Promise(resolve => setTimeout(resolve, 10000)); // Wait for 10 seconds
      await frame.waitForSelector('#bexport', { visible: true });
      context.log("Export button is visible");

      // Click on the "Export" button
      await frame.click('#bexport');
      context.log("Clicked on 'Export' button");

      // Add a wait period to ensure the download has started and completed
      await new Promise(resolve => setTimeout(resolve, 30000)); // Wait for 30 seconds
      context.log("Waited for 30 seconds after clicking 'Export' button");

      // Capture screenshot and HTML after clicking Export button
      await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, 'export', context);

      // Find the downloaded file
      const files = fs.readdirSync(downloadPath);
      const downloadedFile = files.find(file => file.endsWith('.xlsx'));
      if (!downloadedFile) throw new Error('Downloaded file not found');
      const filePath = path.join(downloadPath, downloadedFile);
      context.log("Downloaded file found:", downloadedFile);

      // Save the downloaded file to Azure Blob Storage
      const fileBlobName = `inbound/eSuite/employeeProject.xlsx`;
      await uploadFileToBlob(filePath, blobServiceClient, containerClient, fileBlobName, context);

      // Convert the downloaded file to CSV
      const csvDirPath = path.join('/tmp', `csv_${utcNow}`);
      fs.mkdirSync(csvDirPath, { recursive: true });
      await convertXlsxToCsv(filePath, csvDirPath);

      // Upload the CSV files to Azure Blob Storage
      const csvFiles = fs.readdirSync(csvDirPath);
      for (const csvFile of csvFiles) {
        const csvFilePath = path.join(csvDirPath, csvFile);
        const csvBlobName = `inbound/eSuite/employeeProject.csv`;
        await uploadFileToBlob(csvFilePath, blobServiceClient, containerClient, csvBlobName, context);
      }

    } catch (error) {
      context.log("Error during function execution:", error.message);
      context.res = {
        status: 400,
        body: `Error during processing: ${error.message}`,
        headers: {
          'Content-Type': 'text/plain'
        }
      };
      return;
    } finally {
      await browser.close();
      context.log("Browser closed");
    }

    context.res = {
      status: 200,
      body: "Success!",
      headers: {
        'Content-Type': 'text/plain'
      }
    };
  }
});
