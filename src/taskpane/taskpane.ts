/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async context => {
      // do nothing
    });
  } catch (error) {
    console.error(error);
  }
}
