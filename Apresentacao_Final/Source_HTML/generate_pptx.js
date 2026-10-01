const puppeteer = require('puppeteer');
const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');

async function createPptx() {
  const browser = await puppeteer.launch();
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  
  for (let i = 1; i <= 18; i++) {
    console.log(`Processing page ${i}...`);
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 720, deviceScaleFactor: 2 });
    const fileUrl = 'file://' + path.resolve(__dirname, `pagina_${i}.html`);
    
    await page.goto(fileUrl, { waitUntil: 'networkidle0' });
    
    // Pequeno delay para garantir que as fontes carreguem completamente
    await new Promise(resolve => setTimeout(resolve, 500));
    
    const screenshotPath = path.resolve(__dirname, `slide_${i}.png`);
    await page.screenshot({ path: screenshotPath });
    await page.close();
    
    const slide = pptx.addSlide();
    slide.addImage({ path: screenshotPath, x: 0, y: 0, w: '100%', h: '100%' });
  }
  
  await browser.close();
  
  const outputPath = path.resolve(__dirname, 'Apresentacao.pptx');
  console.log('Saving PPTX to', outputPath);
  await pptx.writeFile({ fileName: outputPath });
  console.log('Cleaning up images...');
  
  for (let i = 1; i <= 18; i++) {
    const screenshotPath = path.resolve(__dirname, `slide_${i}.png`);
    if (fs.existsSync(screenshotPath)) {
        fs.unlinkSync(screenshotPath);
    }
  }
  
  console.log('Done!');
}

createPptx().catch(console.error);
