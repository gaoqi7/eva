const puppeteer = require("puppeteer");

(async () => {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
  await page.setDefaultNavigationTimeout(0);
  await page.goto("http://wbclsb04prod.evaair.com/hostfin/main.asp");
  await page.bringToFront();
  // await page.keyboard.press(String.fromCharCode(13));
  await page.screenshot({ path: "example.png" });

  await browser.close();
})();
