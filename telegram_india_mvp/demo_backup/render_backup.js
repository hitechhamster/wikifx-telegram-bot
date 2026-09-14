const { chromium } = require('playwright');
const path = require('path');

(async () => {
  const browser = await chromium.launch({
    headless: true,
    executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  });
  const page = await browser.newPage({ viewport: { width: 1100, height: 760 }, deviceScaleFactor: 1 });
  await page.goto(`file:///${path.join(__dirname, 'index.html').replace(/\\/g, '/')}`);
  for (const name of ['welcome', 'query', 'news', 'recall']) {
    await page.locator(`#${name}`).screenshot({ path: path.join(__dirname, `${name}.png`) });
  }
  await browser.close();
})();
