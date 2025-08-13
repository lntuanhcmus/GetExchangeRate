const express = require('express');
const puppeteer = require('puppeteer');
const swaggerUi = require('swagger-ui-express');
const swaggerJsdoc = require('swagger-jsdoc');
const XLSX = require('xlsx');

const app = express();
app.use(express.json());

// Cấu hình swagger-jsdoc
const swaggerOptions = {
  definition: {
    openapi: '3.0.0',
    info: {
      title: 'TPBank Exchange API',
      version: '1.0.0',
      description: 'API lấy tỷ giá ngoại tệ TPBank',
    },
  },
  apis: ['./server.js'], // Đường dẫn tới file chứa mô tả API
};

const swaggerSpec = swaggerJsdoc(swaggerOptions);

// Đường dẫn Swagger UI

app.use('/swagger', swaggerUi.serve, swaggerUi.setup(swaggerSpec));
/**
 * @openapi
 * /tygia:
 *   get:
 *     summary: Lấy tỷ giá ngoại tệ theo ngày
 *     parameters:
 *       - in: query
 *         name: date
 *         schema:
 *           type: string
 *           example: "2024-06-01"
 *         required: true
 *         description: Ngày cần lấy tỷ giá (yyyy-MM-dd)
 *     responses:
 *       200:
 *         description: Danh sách tỷ giá
 *         content:
 *           application/json:
 *             schema:
 *               type: array
 *               items:
 *                 type: object
 *                 properties:
 *                   CurrencyCode:
 *                     type: string
 *                   CurrencyName:
 *                     type: string
 *                   BuyCash:
 *                     type: string
 *                   BuyTransfer:
 *                     type: string
 *                   SellCash:
 *                     type: string
 *                   SellTransfer:
 *                     type: string
 */
app.get('/tygia', async (req, res) => {
    const date = req.query.date; // Ngày dạng 'YYYY-MM-DD'
     if (!date) return res.status(400).json({ error: 'Missing date parameter' });
    const formattedDate = formatDate(date); // Định dạng lại ngày nếu cần
    const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
});
    const page = await browser.newPage();
    await page.goto('https://tpb.vn/cong-cu-tinh-toan/ty-gia-ngoai-te', { waitUntil: 'networkidle2' });

    // Chờ và nhập ngày (giả sử có input với id 'datepickerInput')
    await page.waitForSelector('#datepickerInput');
    await page.type('#datepickerInput', formattedDate);

    // Click nút tra cứu (giả sử có id 'search-btn')
    await page.click('#xem-ty-gia');
    await new Promise(r => setTimeout(r, 3000));

    // Lấy dữ liệu bảng
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

    await browser.close();
    res.json(data);
});

/**
 * @openapi
 * /export-excel:
 *   get:
 *     summary: Xuất tỷ giá ngoại tệ ra file Excel theo ngày
 *     parameters:
 *       - in: query
 *         name: date
 *         schema:
 *           type: string
 *           example: "2024-06-01"
 *         required: true
 *         description: Ngày cần lấy tỷ giá (yyyy-MM-dd)
 *     responses:
 *       200:
 *         description: File Excel chứa tỷ giá
 *         content:
 *           application/vnd.openxmlformats-officedocument.spreadsheetml.sheet:
 *             schema:
 *               type: string
 *               format: binary
 */
app.get('/export-excel', async (req, res) => {
    const date = req.query.date;
    if (!date) return res.status(400).json({ error: 'Missing date parameter' });
    const formattedDate = formatDate(date);

    puppeteer.launch({
  headless: true,
  args: ['--no-sandbox', '--disable-setuid-sandbox']
});
    const page = await browser.newPage();
    await page.goto('https://tpb.vn/cong-cu-tinh-toan/ty-gia-ngoai-te', { waitUntil: 'networkidle2' });

    await page.waitForSelector('#datepickerInput');
    await page.type('#datepickerInput', formattedDate);

    await page.click('#xem-ty-gia');
    await new Promise(r => setTimeout(r, 3000));

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

    await browser.close();

    // Tạo file Excel
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
});

function formatDate(dateStr) {
    // dateStr dạng "2025-08-13"
    const [year, month, day] = dateStr.split('-');
    return `${day}/${month}/${year}`;
}

app.listen(3000, () => console.log('Server running on port 3000'));