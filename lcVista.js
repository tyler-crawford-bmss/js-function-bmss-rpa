const puppeteer = require("puppeteer-extra");
const { app } = require("@azure/functions");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const { BlobServiceClient } = require('@azure/storage-blob');
const fs = require('fs');
const path = require('path');
const csv = require('csv-parser'); // Ensure you have the csv-parser package installed


puppeteer.use(StealthPlugin());

function delay(time) {
    return new Promise(resolve => setTimeout(resolve, time));
  };

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

async function uploadFileToBlob(filePath, blobServiceClient, containerClient, blobName, context) {
  const fileBuffer = fs.readFileSync(filePath);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);
  await blockBlobClient.upload(fileBuffer, fileBuffer.length);
  context.log(`File uploaded to Azure Blob Storage at location: ${blobName}`);
}

// Function to convert CSV to JSON and upload to Blob Storage
async function convertCsvToJsonAndUpload(csvFilePath, blobServiceClient, containerClient, blobNameJson, context, jurisdiction) {
    try {
        const results = [];
        return new Promise((resolve, reject) => {
            fs.createReadStream(csvFilePath)
                .pipe(csv())
                .on('data', (data) => {
                    data["License Jurisdiction"] = jurisdiction;
                    results.push(data);
                })
                .on('end', async () => {
                    const jsonData = JSON.stringify(results, null, 4);
                    const blockBlobClient = containerClient.getBlockBlobClient(blobNameJson);
                    await blockBlobClient.upload(jsonData, Buffer.byteLength(jsonData));
                    context.log(`JSON data uploaded to Azure Blob Storage at location: ${blobNameJson}`);
                    resolve(results);
                })
                .on('error', (error) => {
                    context.log(`Error during CSV to JSON conversion and upload: ${error.message}`);
                    reject(error);
                });
        });
    } catch (error) {
        context.log(`Error during CSV to JSON conversion and upload: ${error.message}`);
        throw error;
    }
}

async function processReport(reportName, csvBlobName, jsonBlobName, jurisdiction, page, downloadPath, blobServiceClient, containerClient, context) {
    const reportRowIndex = await page.evaluate((name) => {
        const rows = Array.from(document.querySelectorAll('table.table-default tbody tr'));
        return rows.findIndex(row => {
            const nameCell = row.querySelector('td:first-child');
            return nameCell && nameCell.textContent.trim().includes(name);
        });
    }, reportName);

    if (reportRowIndex !== -1) {
        const dropdownButtonSelector = `table.table-default tbody tr:nth-child(${reportRowIndex + 1}) .test--row-action-button`;
        await page.click(dropdownButtonSelector);
        context.log(`Clicked the dropdown menu for '${reportName}'`);
        await delay(3000);

        const exportToCSV = await page.evaluateHandle(() => {
            const buttons = Array.from(document.querySelectorAll('button[name="format"][value="csv"]'));
            return buttons.find(button => button.textContent.includes('Export to CSV'));
        });

        if (exportToCSV) {
            await exportToCSV.click();
            context.log(`Clicked the 'Export to CSV' button for '${reportName}'`);
        } else {
            throw new Error(`Export to CSV option not found for '${reportName}'`);
        }

        await delay(10000);

        const files = fs.readdirSync(downloadPath);
        const csvFile = files.find(file => file.includes(jurisdiction) && file.endsWith('.csv'));
        if (csvFile) {
            const filePath = path.join(downloadPath, csvFile);

            await uploadFileToBlob(filePath, blobServiceClient, containerClient, csvBlobName, context);

            const jsonArray = await convertCsvToJsonAndUpload(filePath, blobServiceClient, containerClient, jsonBlobName, context, jurisdiction);

            fs.unlinkSync(filePath);
            context.log(`Downloaded CSV file '${csvFile}' uploaded as '${csvBlobName}' and converted to JSON '${jsonBlobName}'`);

            return jsonArray;
        } else {
            throw new Error(`CSV file not found in the download directory for '${jurisdiction}'`);
        }
    } else {
        throw new Error(`Row for '${reportName}' not found`);
    }
}

app.http('lcVista', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    context.log("Function 'lcVista' started.");

    const url = "https://bmssu.lcvista.com/";
    const username = process.env.LCVISTA_USERNAME;
    const password = process.env.LCVISTA_PASSWORD;

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

    let step = 'initial';

    try {
      const page = await browser.newPage();
      await page._client().send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath });
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
      
      // Process Alabama Report
        const alJsonArray = await processReport(
            'Compliance Progress - Alabama', 
            'lcVista/cpe_al.csv', 
            'lcVista/cpe_al.json', 
            'Alabama', 
            page, 
            downloadPath, 
            blobServiceClient, 
            containerClient, 
            context
        );
    
        // Process Mississippi Report
        const msJsonArray = await processReport(
            'Compliance Progress - Mississippi', 
            'lcVista/cpe_ms.csv', 
            'lcVista/cpe_ms.json', 
            'Mississippi', 
            page, 
            downloadPath, 
            blobServiceClient, 
            containerClient, 
            context
        );
    
        // Combine JSON arrays
        const combinedJsonArray = [...alJsonArray, ...msJsonArray];
    
        // Save the combined JSON data to a file and upload it to Blob Storage
        const combinedJsonFilePath = path.resolve('/tmp', 'cpe_compliance.json');
        fs.writeFileSync(combinedJsonFilePath, JSON.stringify(combinedJsonArray, null, 4));
    
        const combinedBlobName = 'lcVista/cpe_compliance.json';
        await uploadFileToBlob(combinedJsonFilePath, blobServiceClient, containerClient, combinedBlobName, context);
    
        context.log(`Combined JSON file uploaded as '${combinedBlobName}'`);
    } catch (error) {
        context.log("Error during function execution:", error.message);
    } finally {
      await browser.close();
      context.log("Browser closed");
    }

    context.res = {
      status: 200,
      body: "Process completed, file uploaded to Blob Storage.",
      headers: {
        'Content-Type': 'text/plain'
      }
    };
  }
});
