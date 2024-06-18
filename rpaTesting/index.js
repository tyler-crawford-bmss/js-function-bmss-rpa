const puppeteer = require('puppeteer');

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');

    console.log('Launching browser...');
    const browser = await puppeteer.launch();
    console.log('Browser launched.');

    console.log('Opening new page...');
    const page = await browser.newPage();
    console.log('New page opened.');

    const url = 'https://github.com';
    console.log(`Navigating to ${url}...`);  // Corrected line
    await page.goto(url);
    console.log(`Navigation to ${url} complete.`);

    const screenshotPath = '/home/site/wwwroot/screenshots/github.png';
    console.log(`Taking screenshot and saving to ${screenshotPath}...`);
    await page.screenshot({ path: screenshotPath });
    console.log(`Screenshot saved to ${screenshotPath}.`);

    console.log('Closing browser...');
    await browser.close();
    console.log('Browser closed.');

    context.res = {
        status: 200,
        body: "Puppeteer script executed successfully."
    };
};
