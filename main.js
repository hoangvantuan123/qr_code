const express = require('express');
const xlsx = require('xlsx');
const puppeteer = require('puppeteer');
const app = express();
const port = 3112;

// Read Excel file
const workbook = xlsx.readFile('BAO_DUONG_MAY_HUNGLONG1_1.xlsx'); // Replace with your Excel file name
const sheetName = workbook.SheetNames[0]; // Get the first sheet
const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]); // Convert sheet to JSON

console.log('sheetData', sheetData);

// Route to render list
app.get('/', (req, res) => {
    let html = '<html><head><title>Danh sách QR</title></head><body>';
    html += '<h1>Danh sách QR Code</h1>';

    // Add Download PDF button
    html += `
        <button onclick="window.location.href='/download-pdf'" 
                style="display: inline-block; margin-bottom: 20px; padding: 10px 20px; background-color: #007bff; color: white; border: none; border-radius: 5px; cursor: pointer;">
            Tải PDF
        </button>
    `;

    // Format table with two columns: Company + Item Code, QR Code
    html += '<div style="display: flex; flex-direction: column; align-items: center;">';

    sheetData.forEach((row) => {
        const company = row['__EMPTY'] ? row['__EMPTY'].split('\n')[0] : 'N/A'; // Company name
        const itemCode = row['id'] || 'N/A'; // Item Code
        const LINK_QR = row['LINK_QR'] || 'N/A'; // QR Code link
        const qrUrl = `https://velvety-decoder-416402.uc.r.appspot.com/api/qrcode2?url=${encodeURIComponent(LINK_QR)}`;

        // Display each item like a "tag" with company name, item code, and QR code
        html += `
            <div style="display: flex; width: 100%; padding: 15px; align-items: center; text-align: center; border: 1px solid #ddd;">
                <!-- QR Code -->
                <div>
                    <img src="${qrUrl}" alt="QR Code" style="width: 120px; height: 120px;">
                </div>

                <!-- Company and item code info -->
                <div style="display: flex; flex-direction: column; align-items: start; text-align: left; margin-left: 15px;">
                    <div style="font-weight: bold; font-size: 16px; margin-bottom: 5px;">${company}</div>
                    <div style="font-size: 14px; margin-bottom: 10px;">ITEM CODE: ${itemCode}</div>
                </div>
            </div>
        `;
    });

    html += '</div></body></html>';
    res.send(html);
});

// Route to download the current view as a PDF
app.get('/download-pdf', async (req, res) => {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    // Generate HTML content for the PDF (same structure as the main route)
    let html = '<html><head><title>Danh sách QR</title></head><body>';
    html += '<h1>Danh sách QR Code</h1>';

    html += `
        <button onclick="window.location.href='/download-pdf'" 
                style="display: inline-block; margin-bottom: 20px; padding: 10px 20px; background-color: #007bff; color: white; border: none; border-radius: 5px; cursor: pointer;">
            Tải PDF
        </button>
    `;

    // Format table with two columns: Company + Item Code, QR Code
    html += '<div style="display: flex; flex-direction: column; align-items: center;">';

    sheetData.forEach((row) => {
        const company = row['__EMPTY'] ? row['__EMPTY'].split('\n')[0] : 'N/A'; // Company name
        const itemCode = row['id'] || 'N/A'; // Item Code
        const LINK_QR = row['LINK_QR'] || 'N/A'; // QR Code link
        const qrUrl = `https://velvety-decoder-416402.uc.r.appspot.com/api/qrcode2?url=${encodeURIComponent(LINK_QR)}`;

        // Display each item like a "tag" with company name, item code, and QR code
        html += `
            <div style="display: flex; width: 100%; padding: 15px; align-items: center; text-align: center; border: 1px solid #ddd;">
                <div>
                    <img src="${qrUrl}" alt="QR Code" style="width: 120px; height: 120px;">
                </div>
                <div style="display: flex; flex-direction: column; align-items: start; text-align: left; margin-left: 15px;">
                    <div style="font-weight: bold; font-size: 16px; margin-bottom: 5px;">${company}</div>
                    <div style="font-size: 14px; margin-bottom: 10px;">ITEM CODE: ${itemCode}</div>
                </div>
            </div>
        `;
    });

    html += '</div></body></html>';

    // Set content for Puppeteer page
    await page.setContent(html);

    // Generate the PDF
    await page.pdf({
        path: 'danh_sach_qr.pdf',
        format: 'A4',
        printBackground: true
    });

    await browser.close();

    // Send a response to the user to download the file
    res.download('danh_sach_qr.pdf', 'danh_sach_qr.pdf', (err) => {
        if (err) {
            console.error('Error while downloading the file:', err);
        } else {
            console.log('PDF file has been downloaded successfully');
        }
    });
});

// Start server
app.listen(port, () => {
    console.log(`Server is running at http://localhost:${port}`);
});
