const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const https = require('https')
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

class Payload {
  constructor(actor, input) {
      this.actor = actor;
      this.input = input;
  }
}

function sendRequest(url) {
  return new Promise((resolve, reject) => {
    const host = "api.scrapeless.com";
    const urlScrapper = `https://${host}/api/v1/scraper/request`;
    const token = "sk_ndQZCLNmJT73I1OMbsEqQ4pmqFsZMhB71usbNBz8kJQfqPAoLrdlUziM1p2uQRrI";

    const inputData = {
      url,
    };

    const payload = new Payload("scraper.shopee", inputData);
    const jsonPayload = JSON.stringify(payload);

    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-token': token,
      },
    };

    const req = https.request(urlScrapper, options, (res) => {
      let body = '';

      res.on('data', (chunk) => {
        body += chunk;
      });

      res.on('end', () => {
        try {
          resolve(body); // Resolve with raw response body
        } catch (error) {
          reject(new Error('Failed to process response: ' + error.message));
        }
      });
    });

    req.on('error', (error) => {
      reject(new Error('Request failed: ' + error.message));
    });

    req.write(jsonPayload);
    req.end();
  });
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
  const outputPath = path.join(path.dirname(filePath), `results_${Date.now()}.xlsx`)
  
  const inputWorkbook = new ExcelJS.Workbook()
  await inputWorkbook.xlsx.readFile(filePath)
  const inputSheet = inputWorkbook.worksheets[0]

  const outputWorkbook = new ExcelJS.Workbook()
  const outputSheet = outputWorkbook.addWorksheet('Prices')
  outputSheet.addRow(['URL', 'Price', 'Status'])

  for (let i = 2; i <= inputSheet.rowCount; i++) {
    const url = inputSheet.getCell(`K${i}`).value
    if (!url || !url.hyperlink) continue

    try {
      console.log(`Processing URL: #${i}`);
      const price = await getPrice(url.hyperlink)
      outputSheet.addRow([url.hyperlink, price, 'Success'])
    } catch (error) {
      outputSheet.addRow([url.hyperlink, 'N/A', `Error: ${error.message}`])
      console.log(error.message);
    }
  }

  await outputWorkbook.xlsx.writeFile(outputPath)
  return outputPath
})

async function getPrice(url) {
  try {
    const response = await sendRequest(url)
    const data = JSON.parse(response)
    
    // Check for error in response
    if (data.error) {
      throw new Error(`API error ${data.error}: ${data.tracking_id || 'No tracking ID'}`)
    }

    // Extract price (adjust path based on actual API response structure)
    // Assuming price is in data.result.price or similar
    if(!data.data){
      console.log(data);
      throw new Error('Not My error, scrapeless API error');
    }

    const singleValue = data.data.data.product_price.price.single_value || 0;
    const rangeMin = data.data.data.product_price.price.range_min || 'N/A';

    // Use rangeMin if singleValue is less than or equal to 0
    const price = singleValue > 0 ? singleValue : rangeMin;

    if (price === 'N/A') {
      throw new Error('Price not found in response')
    }
    
    // Format price (assuming it's in smallest unit, e.g., cents or VND without decimals)
    return (price / 100000) // Adjust divisor based on currency (e.g., 100 for VND)
  } catch (error) {
    throw new Error(`Failed to get price: ${error.message}`)
  }
}

app.whenReady().then(createWindow)