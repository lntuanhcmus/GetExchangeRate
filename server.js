const express = require('express');
const puppeteer = require('puppeteer');
const swaggerUi = require('swagger-ui-express');
const swaggerJsdoc = require('swagger-jsdoc');
const XLSX = require('xlsx');

const app = express();
app.use(express.json());

const swaggerOptions = {
  definition: {
    openapi: '3.0.0',
    info: {
      title: 'TPBank Exchange API',
      version: '1.0.0',
      description: 'API lấy tỷ giá ngoại tệ TPBank',
    },
  },
  apis: ['./server.js'],
};

const swaggerSpec = swaggerJsdoc(swaggerOptions);
app.use('/swagger', swaggerUi.serve, swaggerUi.setup(swaggerSpec));

app.get('/tygia', async (req, res) => {
    const date = req.query.date;
    if (!date) return res.status(400).json({ error: 'Missing date parameter' });
    const formattedDate = formatDate(date);

    let browser;
    try {
        browser = await puppeteer.launch({
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        });
        const page = await browser.newPage();
        await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36');
        await page.goto('https://tpb.vn/cong-cu-tinh-toan/ty-gia-ngoai-te', { waitUntil: 'networkidle2' });

        await page.waitForSelector('#datepickerInput', { timeout: 300000 });
        await page.type('#datepickerInput', formattedDate);

        await page.click('#xem-ty-gia');
        await page.waitForSelector('.table', { timeout: 300000 });

        const data = await page.evaluate(() => {
            const table = document.querySelector('.table');
            if (!table) return [];
            const rows = Array.from(table.querySelectorAll('tbody tr'));
            return rows.map(row => {
                const cols = row.querySelectorAll('td');
                return {
                    CurrencyCode: cols[0]?.innerText.trim(),
                    CurrencyName: cols[1]?.innerText.trim(),
                    BuyCash: cols[2]?.innerText.trim().replace(/,/g, ''),
                    BuyTransfer: cols[3]?.innerText.trim().replace(/,/g, ''),
                    SellCash: cols[4]?.innerText.trim().replace(/,/g, ''),
                    SellTransfer: cols[5]?.innerText.trim().replace(/,/g, '')
                };
            });
        });

        res.json(data);
    } catch (error) {
        res.status(500).json({ error: error.message });
    } finally {
        if (browser) await browser.close();
    }
});

app.get('/export-excel', async (req, res) => {
    const date = req.query.date;
    if (!date) return res.status(400).json({ error: 'Missing date parameter' });
    const formattedDate = formatDate(date);

    let browser;
    try {
        browser = await puppeteer.launch({
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        });
        const page = await browser.newPage();
        await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36');
        await page.goto('https://tpb.vn/cong-cu-tinh-toan/ty-gia-ngoai-te', { waitUntil: 'networkidle2' });

        await page.waitForSelector('#datepickerInput', { timeout: 300000 });
        await page.type('#datepickerInput', formattedDate);

        await page.click('#xem-ty-gia');
        await page.waitForSelector('.table', { timeout: 300000 });

        const data = await page.evaluate(() => {
            const table = document.querySelector('.table');
            if (!table) return [];
            const rows = Array.from(table.querySelectorAll('tbody tr'));
            return rows.map(row => {
                const cols = row.querySelectorAll('td');
                return {
                    CurrencyCode: cols[0]?.innerText.trim(),
                    CurrencyName: cols[1]?.innerText.trim(),
                    BuyCash: cols[2]?.innerText.trim().replace(/,/g, ''),
                    BuyTransfer: cols[3]?.innerText.trim().replace(/,/g, ''),
                    SellCash: cols[4]?.innerText.trim().replace(/,/g, ''),
                    SellTransfer: cols[5]?.innerText.trim().replace(/,/g, '')
                };
            });
        });

        const wsData = [
            ['CurrencyCode', 'CurrencyName', 'BuyCash', 'BuyTransfer', 'SellCash', 'SellTransfer'],
            ...data.map(item => [
                item.CurrencyCode,
                item.CurrencyName,
                item.BuyCash,
                item.BuyTransfer,
                item.SellCash,
                item.SellTransfer
            ])
        ];
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, 'ExchangeRates');
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        res.setHeader('Content-Disposition', `attachment; filename="ExchangeRates_${date}.xlsx"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (error) {
        res.status(500).json({ error: error.message });
    } finally {
        if (browser) await browser.close();
    }
});

function formatDate(dateStr) {
    const [year, month, day] = dateStr.split('-');
    return `${day}/${month}/${year}`;
}

app.listen(3000, () => console.log('Server running on port 3000'));