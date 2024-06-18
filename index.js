const puppeteer = require('puppeteer');

async function run() {
console.log('Launching browser...');
const browser = await puppeteer.launch();
console.log('Browser launched.');

console.log('Opening new page...');
const page = await browser.newPage();
console.log('New page opened.');

const url = 'https://github.com';
console.log(Navigating to ${url}...);
await page.goto(url);
console.log(Navigation to ${url} complete.);

const screenshotPath = 'screenshots/github.png';
console.log(Taking screenshot and saving to ${screenshotPath}...);
await page.screenshot({ path: screenshotPath });
console.log(Screenshot saved to ${screenshotPath}.);

console.log('Closing browser...');
await browser.close();
console.log('Browser closed.');
}

run().then(() => {
console.log('Script completed successfully.');
}).catch(error => {
console.error('Error during script execution:', error);
});
