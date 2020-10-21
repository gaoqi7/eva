//export PDF file from Financial System
const robot = require("robotjs");
const sleep = require("sleep");
const toPath = "C:\\Users\\H36360\\Documents\\eva\\pdf";
//How many reports need to be exported.
const rpt = process.argv[2];
// var mouse = robot.getMousePos();
// console.log("Mouse is at x:" + mouse.x + " y:" + mouse.y)
const startPoint = [93, 299];
function pdfExport(times) {
  sleep.sleep(5);
  let originalY = startPoint[1];
  for (i = 0; i < times; i++) {
    originalY = startPoint[1] + i * 33;
    const pdfFileName = `\\${times - i}.pdf`;
    robot.moveMouse(93, originalY);
    robot.mouseClick();
    sleep.sleep(3);

    robot.moveMouse(555, 115);
    robot.mouseClick();
    robot.typeString("a");
    robot.keyTap("tab");
    robot.keyTap("tab");
    robot.keyTap("tab");
    robot.keyTap("tab");
    robot.typeString(`${toPath}${pdfFileName}`);
    robot.keyTap("tab");
    robot.keyTap("enter");
    sleep.sleep(3);

    robot.keyTap("tab");
    sleep.sleep(3);
    robot.keyTap("enter");
    sleep.sleep(3);
  }
  console.log(`Done! Successfully export ${rpt} pdf files.`);
}

pdfExport(rpt);
