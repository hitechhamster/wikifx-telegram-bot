const { chromium } = require('playwright');
const path = require('path');

(async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({
    viewport: { width: 2600, height: 1500 },
    deviceScaleFactor: 1.5
  });

  const source = path.resolve('mermaid-process-render.html');
  await page.goto(`file:///${source.replace(/\\/g, '/')}`, { waitUntil: 'networkidle' });
  await page.waitForSelector('.mermaid svg', { state: 'visible', timeout: 30000 });
  await page.evaluate(async () => {
    await document.fonts.ready;
  });

  const area = page.locator('#export-area');
  await area.screenshot({
    path: path.resolve('自然北极星私域运营闭环图.png'),
    type: 'png',
    animations: 'disabled'
  });

  await browser.close();
})();
