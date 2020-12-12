import puppeteer from "puppeteer";
import cron from "cron";
import fetch from "node-fetch"
function report() {
  puppeteer.launch({ headless:true,slowMo:250,args:['--no-sandbox','--disable-setuid-sandbox']}).then(async (browser) => {
    const page = await browser.newPage();
    await page.goto(
      "https://docs.google.com/forms/d/e/1FAIpQLScK0Wzq_ti8cF9LSNZ3wfvaTTlbJTe0xFGwT7Ij1VcjTCHB8g/viewform",
      { waitUntil: "networkidle2" }
    );
    await page.bringToFront();
    await page.keyboard.press("Tab");
    await page.keyboard.type("H36360", {
      delay: 100,
    });
    await page.keyboard.press("Tab");
    await page.keyboard.type("Rick Gao");
    await page.keyboard.press("Tab");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Tab");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Tab");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Tab");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Tab");
    await page.keyboard.press("Tab");
    await page.keyboard.press("Enter");

    await page.screenshot({ path: "hr.png" });
    await browser.close();
  });
	notify()
}
function notify(){
	const d = new Date()
	const t = `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()} 07:15`
	const url = `https://maker.ifttt.com/trigger/health_report/with/key/Rr_JZE2LHCLfHwxW52u_a`
	const body = {value1:t}
	console.log(body)
	fetch(url,{method:'POST',
	           body:JSON.stringify(body),
       	           headers: { 'Content-Type': 'application/json' },
	}
	)
}


const CronJob = cron.CronJob;
const job = new CronJob(`15 7 * * *`, function () {
	  report();
	});
job.start();

