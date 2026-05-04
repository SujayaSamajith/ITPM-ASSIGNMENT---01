const { chromium } = require('@playwright/test');
(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto('https://www.pixelssuite.com/chat-translator');
  await page.waitForTimeout(2000); // Wait for React to render
  const html = await page.content();
  console.log(html);
  await browser.close();
})();
