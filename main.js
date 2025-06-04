const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const puppeteer = require('puppeteer')
const ExcelJS = require('exceljs')

let mainWindow

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    }
  })

  mainWindow.loadFile('index.html')
}

// File dialog handler
ipcMain.handle('open-file-dialog', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  })
  return result.filePaths[0]
})

// Excel processing handler
ipcMain.handle('process-excel', async (_, filePath) => {
  const browser = await puppeteer.launch({ headless: true })
  const outputPath = path.join(path.dirname(filePath), `results_${Date.now()}.xlsx`)
  
  const inputWorkbook = new ExcelJS.Workbook()
  await inputWorkbook.xlsx.readFile(filePath)
  const inputSheet = inputWorkbook.worksheets[0]

  const outputWorkbook = new ExcelJS.Workbook()
  const outputSheet = outputWorkbook.addWorksheet('Prices')
  outputSheet.addRow(['URL', 'Price', 'Status'])

  for (let i = 2; i <= inputSheet.rowCount; i++) {
    const url = inputSheet.getCell(`K${i}`).value
    if (!url) continue

    try {
      const price = await getPrice(browser, url)
      outputSheet.addRow([url, price, 'Success'])
    } catch (error) {
      outputSheet.addRow([url, 'N/A', `Error: ${error.message}`])
      console.log(error.message);
      
    }
  }

  await outputWorkbook.xlsx.writeFile(outputPath)
  await browser.close()
  return outputPath
})

async function getPrice(browser, url) {
  const page = await browser.newPage()
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
  
  try {
    await page.goto(url.hyperlink, { waitUntil: 'networkidle2', timeout: 30000 })
    await page.waitForSelector('div[class*="IZPeQz B67UQ0"]', { timeout: 300000 })
    return await page.$eval('div[class*="IZPeQz B67UQ0"]', el => el.textContent.trim())
  } finally {
    await page.close()
  }
}

app.whenReady().then(createWindow)