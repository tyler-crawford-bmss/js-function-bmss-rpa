const puppeteer = require("puppeteer-extra");
const { app } = require("@azure/functions");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const { BlobServiceClient } = require('@azure/storage-blob');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

puppeteer.use(StealthPlugin());

// Function to capture screenshots and HTML content
async function captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, step, context) {
  try {
    const timeStamp = Date.now();
    const screenshotBuffer = await page.screenshot({ fullPage: true });
    
    let htmlContent = await page.content();

    const frames = page.frames();
    for (const frame of frames) {
      if (frame.parentFrame() !== null) {
        const frameName = frame.name();
        const frameContent = await frame.content();
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

async function runLongProcess(context, url, username, password, utcNow, containerClient, downloadPath) {
  const browser = await puppeteer.launch({
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    headless: true
  });

  try {
    const existingCsvFileName = 'inbound/eSuite/employeeProject.csv';
    const lastWeekCsvFileName = 'inbound/eSuite/employeeProject_lastWeek.csv';
    await copyBlob(blobServiceClient, process.env.AZURE_CONTAINER_NAME, existingCsvFileName, process.env.AZURE_CONTAINER_NAME, lastWeekCsvFileName, context);
    
    context.log("Existing CSV file saved as last week's file");  
    let page;

    page = await browser.newPage();

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

    await new Promise(resolve => setTimeout(resolve, 30000));
    context.log("Waited for 30 seconds after login");

    await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, 'login', context);

    await page.click('#ahrefreports_menu');
    context.log("Clicked on 'Reports' tab");

    await new Promise(resolve => setTimeout(resolve, 30000));
    context.log("Waited for 30 seconds after clicking 'Reports' tab");

    let frame = page.frames().find(frame => frame.name() === 'main');
    if (!frame) throw new Error('Frame "main" not found');

    context.log("Switched to frame 'main'");

    const newPagePromise = new Promise(resolve => {
      page.once('popup', popup => resolve(popup));
    });

    await frame.click('a[href="javascript:OnEmployeeSearch()"] img[src="binoculars.gif"]');
    context.log('Clicked on the "binoculars.gif" button');

    const newPage = await newPagePromise;

    context.log('New page opened');

    await new Promise(resolve => setTimeout(resolve, 30000));
    context.log('Waited for 30 seconds after opening employee search page');

    await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'employee_search', context);

    const resultsFrame = newPage.frames().find(frame => frame.name() === 'results');
    if (!resultsFrame) throw new Error('Frame "results" not found');
    context.log('Accessed the "results" frame');

    await resultsFrame.click('tr.list-h input[type="checkbox"][onclick="return onSelectAll(this)"]');
    context.log('Step 1: Checked the header checkbox');
    await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step1', context);
    await new Promise(resolve => setTimeout(resolve, 5000));

    await resultsFrame.click('input[type="button"][value="Remove"][onclick="onRemoveSelected(\'ShowSelectedList\')"]');
    context.log('Step 2: Clicked the "Remove" button');
    await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step2', context);
    await new Promise(resolve => setTimeout(resolve, 5000));

    const criteriaFrame = newPage.frames().find(frame => frame.name() === 'criteria');
    if (!criteriaFrame) throw new Error('Frame "criteria" not found');
    context.log('Accessed the "criteria" frame');

    await criteriaFrame.select('select[name="emplstatus"]', '1');
    context.log('Step 3: Selected "Active" from the dropdown');
    await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step3', context);
    await new Promise(resolve => setTimeout(resolve, 5000));

    await criteriaFrame.click('input[type="button"][value="Search"]');
    context.log('Step 4: Clicked the "Search" button');
    await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step4', context);
    await new Promise(resolve => setTimeout(resolve, 10000));

    await resultsFrame.click('tr.list-h input[type="checkbox"][onclick="return onSelectAll(this)"]');
    context.log('Step 5: Checked the header checkbox again');
    await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step5', context);
    await new Promise(resolve => setTimeout(resolve, 5000));

    await resultsFrame.click('input[type="button"][value="Add"][onclick="onAddSelected()"]');
    context.log('Step 6: Clicked the "Add" button');
    await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step6', context);
    await new Promise(resolve => setTimeout(resolve, 10000));

    const framesAfterAdd = newPage.frames();
    framesAfterAdd.forEach(frame => {
      context.log(`Frame name: ${frame.name()}, URL: ${frame.url()}`);
    });

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
      await captureAndUploadState(newPage, containerClient, sanitizedWebsite, utcNow, 'step7_before_click', context);
      await okButtonFrame.click('input[type="button"][id="doIt"][value="OK"]');
      context.log('Step 7: Clicked the "OK" button');
      await new Promise(resolve => setTimeout(resolve, 10000));
    } else {
      throw new Error('Could not find the "OK" button after clicking "Add"');
    }

    try {
      await newPage.close();
      context.log('Closed the employee search window');
    } catch (e) {
      context.log('The employee search window was already closed');
    }

    await page.bringToFront();
    context.log('Brought main page to the front');

    frame = page.frames().find(frame => frame.name() === 'main');
    if (!frame) throw new Error('Frame "main" not found after closing employee search window');

    context.log("Switched back to frame 'main'");

    await new Promise(resolve => setTimeout(resolve, 10000));
    await frame.waitForSelector('#bexport', { visible: true });
    context.log("Export button is visible");

    await frame.click('#bexport');
    context.log("Clicked on 'Export' button");

    await new Promise(resolve => setTimeout(resolve, 30000));
    context.log("Waited for 30 seconds after clicking 'Export' button");

    await captureAndUploadState(page, containerClient, sanitizedWebsite, utcNow, 'export', context);

    const files = fs.readdirSync(downloadPath);
    const downloadedFile = files.find(file => file.endsWith('.xlsx'));
    if (!downloadedFile) throw new Error('Downloaded file not found');
    const filePath = path.join(downloadPath, downloadedFile);
    context.log("Downloaded file found:", downloadedFile);

    const fileBlobName = `inbound/eSuite/employeeProject.xlsx`;
    await uploadFileToBlob(filePath, blobServiceClient, containerClient, fileBlobName, context);

    const csvDirPath = path.join('/tmp', `csv_${utcNow}`);
    fs.mkdirSync(csvDirPath, { recursive: true });
    await convertXlsxToCsv(filePath, csvDirPath);

    const csvFiles = fs.readdirSync(csvDirPath);
    for (const csvFile of csvFiles) {
      const csvFilePath = path.join(csvDirPath, csvFile);
      const csvBlobName = `inbound/eSuite/employeeProject.csv`;
      await uploadFileToBlob(csvFilePath, blobServiceClient, containerClient, csvBlobName, context);
    }

  } catch (error) {
    context.log("Error during function execution:", error.message);
  } finally {
    await browser.close();
    context.log("Browser closed");
  }
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

    const utcNow = Date.now();
    const sanitizedWebsite = url.replace(/[^a-zA-Z0-9]/g, '_');
    const downloadPath = path.resolve('/tmp', `download_${utcNow}`);
    fs.mkdirSync(downloadPath, { recursive: true });

    runLongProcess(context, url, username, password, utcNow, containerClient, downloadPath);

    context.res = {
      status: 202,
      body: "Processing started. You will be notified upon completion.",
      headers: {
        'Content-Type': 'text/plain'
      }
    };
  }
});
