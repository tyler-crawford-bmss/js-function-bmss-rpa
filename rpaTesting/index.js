const chromium = require('chrome-aws-lambda');

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');

    try {
        context.log('Launching browser...');
        const browser = await chromium.puppeteer.launch({
            args: chromium.args,
            defaultViewport: chromium.defaultViewport,
            executablePath: await chromium.executablePath,
            headless: chromium.headless,
        });
        context.log('Browser launched.');

        context.log('Opening new page...');
        const page = await browser.newPage();
        context.log('New page opened.');

        const url = 'https://github.com';
        context.log(`Navigating to ${url}...`);
        await page.goto(url);
        context.log(`Navigation to ${url} complete.`);

        const screenshotPath = '/home/site/wwwroot/screenshots/github.png';
        context.log(`Taking screenshot and saving to ${screenshotPath}...`);
        await page.screenshot({ path: screenshotPath });
        context.log(`Screenshot saved to ${screenshotPath}.`);

        context.log('Closing browser...');
        await browser.close();
        context.log('Browser closed.');

        context.res = {
            status: 200,
            body: "Puppeteer script executed successfully."
        };
    } catch (error) {
        context.log.error('Error during script execution:', error);
        context.res = {
            status: 500,
            body: `Error during script execution: ${error.message}`
        };
    }
};
