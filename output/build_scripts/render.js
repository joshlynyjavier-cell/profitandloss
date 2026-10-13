const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' }).catch(()=>chromium.launch());
  const p = await b.newPage({ viewport: { width: 1280, height: 720 } });
  await p.goto('file://' + process.argv[2]); await p.waitForTimeout(500);
  await p.pdf({ path: process.argv[3], width: '13.333in', height: '7.5in', printBackground: true });
  // overflow check
  const o = await p.evaluate(() => [...document.querySelectorAll('.slide')].map(s => [s.scrollHeight, s.clientHeight]));
  console.log(JSON.stringify(o));
  await b.close();
})();
