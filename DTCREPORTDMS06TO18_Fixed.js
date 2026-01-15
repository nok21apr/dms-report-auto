const puppeteer = require('puppeteer');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');

// ฟังก์ชันสำหรับแปลงวันที่เป็น YYYY-MM-DD
function getTodayFormatted() {
    const date = new Date();
    const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Bangkok' };
    const thaiDate = new Intl.DateTimeFormat('en-CA', options).format(date);
    return thaiDate;
}

(async () => {
    // --- ส่วนการรับค่าจาก Secrets ---
    const USERNAME = process.env.DTC_USERNAME;
    const PASSWORD = process.env.DTC_PASSWORD;
    const EMAIL_USER = process.env.EMAIL_USER;
    const EMAIL_PASS = process.env.EMAIL_PASS;
    const EMAIL_TO   = process.env.EMAIL_TO;

    if (!USERNAME || !PASSWORD || !EMAIL_USER || !EMAIL_PASS || !EMAIL_TO) {
        console.error('Error: Missing required secrets.');
        process.exit(1);
    }

    console.log('Launching browser...');
    const downloadPath = path.resolve('./downloads');
    if (!fs.existsSync(downloadPath)) {
        fs.mkdirSync(downloadPath);
    }

    const browser = await puppeteer.launch({
        headless: true,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--start-maximized'
        ]
    });
    
    const page = await browser.newPage();
    // เพิ่ม Timeout รวมเป็น 3 นาที (180s) เผื่อเว็บช้ามาก
    const timeout = 180000; 
    page.setDefaultTimeout(timeout);
    
    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', {
        behavior: 'allow',
        downloadPath: downloadPath,
    });

    await page.setViewport({ width: 1920, height: 1080 });

    try {
        // --- 1. Login ---
        console.log('Navigating to login page...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'networkidle0' });

        await page.waitForSelector('#txtname');
        await page.type('#txtname', USERNAME);
        
        await page.keyboard.press('Tab');
        await new Promise(r => setTimeout(r, 500));
        await page.type('#txtpass', PASSWORD);

        console.log('Submitting login...');
        await page.keyboard.press('Enter');

        // --- แก้ไขจุดที่ 1: เพิ่มเวลาการรอโหลดข้อมูลเป็น 60 วินาที ---
        console.log('Waiting 60 seconds for page data to load completely...');
        await new Promise(r => setTimeout(r, 60000));

        // --- แก้ไขจุดที่ 2: เช็คว่า Overlay (ถ้ามี) หายไปแล้ว และ Sidebar พร้อมใช้งาน ---
        console.log('Checking if Dashboard is ready...');
        try {
            // รอให้ Sidebar ปรากฏ
            await page.waitForSelector('#sidebar', { visible: true, timeout: 30000 });
        } catch (e) {
            console.log('Warning: Sidebar taking long to appear...');
        }

        // --- 2. เข้าเมนูรายงาน ---
        console.log('Clicking Report Tab...');
        const reportSelector = '#sidebar li:nth-of-type(5) i';
        
        // รอจนกว่าปุ่มจะคลิกได้ (Clickable)
        await page.waitForSelector(reportSelector, { visible: true });
        
        // ใช้การคลิก 2 แบบเผื่อไว้ (Click ปกติ และ Evaluate Click)
        try {
            await page.click(reportSelector);
        } catch (e) {
            await page.$eval(reportSelector, el => el.click());
        }

        // --- 3. เลือกรายงาน DMS ---
        console.log('Clicking DMS Status Report...');
        await new Promise(r => setTimeout(r, 3000)); // รอเมนูเลื่อนลงมา

        try {
             const dmsReportXPath = "//*[contains(text(), 'รายงานสถานะ DMS')]";
             // ใช้ page.$$ สำหรับ Puppeteer v23+
             const elements = await page.$$(`xpath/${dmsReportXPath}`);
             
             if (elements.length > 0) {
                 // ตรวจสอบว่า element มองเห็นได้ (visible) ก่อนคลิก
                 // (บางทีเมนูยังไม่กางออกมา)
                 await new Promise(r => setTimeout(r, 1000)); 
                 await elements[0].click();
             } else {
                 throw new Error("Link not found");
             }
        } catch (e) {
            console.log("Using fallback selector for DMS link...");
            // Selector สำรอง
            await page.click('div:nth-of-type(5) > div:nth-of-type(2) li:nth-of-type(1) > a');
        }

        await new Promise(r => setTimeout(r, 10000)); // รอหน้า Report โหลด (เพิ่มเป็น 10 วิ)

        // --- 3.5 เลือกช่วงเวลา 06:00 - 18:00 ของวันนี้ ---
        console.log('Setting Date Range: 06:00 - 18:00...');
        const todayStr = getTodayFormatted();
        const startDateTime = `${todayStr} 06:00`;
        const endDateTime = `${todayStr} 18:00`;

        // รอช่องวันที่พร้อม
        await page.waitForSelector('#date9', { visible: true });
        
        await page.evaluate(() => document.getElementById('date9').value = '');
        await page.type('#date9', startDateTime);

        await page.evaluate(() => document.getElementById('date10').value = '');
        await page.type('#date10', endDateTime);
        
        console.log('Clicking Search to update report...');
        try {
            const searchBtnXPath = "//*[contains(text(), 'ค้นหา')] | //span[contains(@class, 'icon-search')] | //i[contains(@class, 'icon-search')]";
            const searchBtns = await page.$$(`xpath/${searchBtnXPath}`);
            
            if (searchBtns.length > 0) {
                await searchBtns[0].click();
            } else {
                await page.click('td:nth-of-type(5) > span');
            }
            // รอข้อมูลโหลดใหม่หลังจากกดค้นหา (สำคัญ)
            console.log('Waiting for report data to update...');
            await new Promise(r => setTimeout(r, 10000)); 
        } catch (e) {
            console.log('Warning: Could not click Search button.', e.message);
        }

        // --- 4. กดปุ่ม Excel (Download) ---
        console.log('Clicking Export/Excel...');
        
        cleanDownloadFolder(downloadPath);

        const excelBtnSelector = '#btnexport, button[title="Excel"], ::-p-aria(Excel)';
        await page.waitForSelector(excelBtnSelector, { visible: true, timeout: 15000 });
        await page.click(excelBtnSelector);
        
        console.log('Waiting for download (20s)...');
        await new Promise(r => setTimeout(r, 20000));

        // --- 5. ส่ง Email และ ลบไฟล์ ---
        console.log('Processing email...');
        
        const recentFile = getMostRecentFile(downloadPath);
        
        if (recentFile) {
            const filePath = path.join(downloadPath, recentFile.file);
            const fileName = recentFile.file;
            const subjectLine = `${fileName} ช่วง0600ถึง1800`;

            await sendEmail({
                user: EMAIL_USER,
                pass: EMAIL_PASS,
                to: EMAIL_TO,
                subject: subjectLine,
                attachmentPath: filePath
            });

            console.log('Deleting downloaded file...');
            try {
                fs.unlinkSync(filePath);
                console.log('File deleted successfully.');
            } catch (err) {
                console.error('Error deleting file:', err);
            }

        } else {
            console.log('No file downloaded to send.');
            throw new Error('Download failed or no file found');
        }

        console.log('Script completed.');

    } catch (error) {
        console.error('Error occurred:', error);
        await page.screenshot({ path: 'error_screenshot.png' });
        process.exit(1);
    } finally {
        await browser.close();
    }
})();

async function sendEmail({ user, pass, to, subject, attachmentPath }) {
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: { user, pass }
    });

    const mailOptions = {
        from: user,
        to: to,
        subject: subject,
        text: 'รายงาน DMS ประจำช่วงเวลา 06:00 - 18:00\n\n(Auto-generated email)',
        attachments: attachmentPath ? [{ path: attachmentPath }] : []
    };

    const info = await transporter.sendMail(mailOptions);
    console.log('Email sent: ' + info.response);
}

const getMostRecentFile = (dir) => {
    try {
        const files = fs.readdirSync(dir);
        const validFiles = files.filter(file => fs.lstatSync(path.join(dir, file)).isFile() && !file.startsWith('.'));
        if (validFiles.length === 0) return null;
        return validFiles
            .map(file => ({ file, mtime: fs.lstatSync(path.join(dir, file)).mtime }))
            .sort((a, b) => b.mtime.getTime() - a.mtime.getTime())[0];
    } catch (e) { return null; }
};

const cleanDownloadFolder = (dir) => {
    try {
        if (fs.existsSync(dir)) {
            const files = fs.readdirSync(dir);
            for (const file of files) {
                fs.unlinkSync(path.join(dir, file));
            }
        }
    } catch (e) {}
};
