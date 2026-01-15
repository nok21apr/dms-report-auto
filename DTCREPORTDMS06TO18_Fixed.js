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
    
    // --- Setup ตาม Code เก่า ---
    // Timeout 5 นาที
    page.setDefaultNavigationTimeout(300000);
    page.setDefaultTimeout(300000);

    await page.emulateTimezone('Asia/Bangkok');
    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });

    await page.setViewport({ width: 1920, height: 1080 });

    try {
        // ---------------------------------------------------------
        // Step 1: Login
        // ---------------------------------------------------------
        console.log('1️⃣ Step 1: Login...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#txtname', { visible: true, timeout: 60000 });
        await page.type('#txtname', USERNAME);
        await page.type('#txtpass', PASSWORD);
        
        console.log('   Clicking Login...');
        await Promise.all([
            page.evaluate(() => {
                const btn = document.getElementById('btnLogin');
                if(btn) btn.click();
            }),
            page.waitForFunction(() => !document.querySelector('#txtname'), { timeout: 60000 })
        ]);
        console.log('✅ Login Success');

        // ---------------------------------------------------------
        // Step 2: Navigate to Report (Direct URL)
        // ---------------------------------------------------------
        console.log('2️⃣ Step 2: Go to Report Page (Direct URL)...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_other_status.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#date9', { visible: true, timeout: 60000 });
        console.log('✅ Report Page Loaded');

        // ---------------------------------------------------------
        // Step 2.5: Select Truck "ทั้งหมด" (Robust Version)
        // ---------------------------------------------------------
        console.log('   Selecting Truck "ทั้งหมด"...');
        await page.waitForSelector('#ddl_truck', { visible: true, timeout: 60000 });

        // รอจนกว่า Option "ทั้งหมด" จะปรากฏจริง (ป้องกัน Dropdown โหลดไม่เสร็จ)
        await page.waitForFunction(() => {
            const select = document.getElementById('ddl_truck');
            if (!select) return false;
            // เช็คว่ามี option ที่มีคำว่า "ทั้งหมด" หรือ "All" หรือไม่
            return Array.from(select.options).some(opt => opt.text.includes('ทั้งหมด') || opt.text.toLowerCase().includes('all'));
        }, { timeout: 60000 });

        // ทำการเลือกจริง
        await page.evaluate(() => {
            var selectElement = document.getElementById('ddl_truck'); 
            if (selectElement) {
                var options = selectElement.options; 
                for (var i = 0; i < options.length; i++) { 
                    if (options[i].text.includes('ทั้งหมด') || options[i].text.toLowerCase().includes('all')) { 
                        selectElement.value = options[i].value; 
                        var event = new Event('change', { bubbles: true });
                        selectElement.dispatchEvent(event);
                        break; 
                    } 
                }
            }
        });
        console.log('✅ Truck "ทั้งหมด" Selected');

        // ---------------------------------------------------------
        // Step 2.6: Select Report Types (3 Items)
        // ---------------------------------------------------------
        console.log('   Selecting 3 Report Types...');
        
        // รายชื่อรายงานที่ต้องการเลือก (ตามที่คุณระบุ)
        const reportTypesToSelect = [
            "แจ้งเตือนมีความง่วงระดับ 1",
            "แจ้งเตือนมีความง่วงระดับ 2",
            "แจ้งเตือนการหาวนอน"
        ];

        for (const typeName of reportTypesToSelect) {
            try {
                // ใช้ XPath แบบ normalize-space เพื่อตัดช่องว่างส่วนเกินออก ทำให้หาเจอแม่นยำขึ้น
                const xpath = `//label[contains(normalize-space(.), '${typeName}')] | //span[contains(normalize-space(.), '${typeName}')]`;
                const elements = await page.$$(`xpath/${xpath}`);
                
                if (elements.length > 0) {
                    const isChecked = await page.evaluate(el => {
                        const input = el.tagName === 'INPUT' ? el : (el.querySelector('input') || document.getElementById(el.getAttribute('for')));
                        return input ? input.checked : false;
                    }, elements[0]);

                    if (!isChecked) {
                        await elements[0].click();
                        console.log(`      Clicked: ${typeName}`);
                    } else {
                        console.log(`      Already checked: ${typeName}`);
                    }
                } else {
                    console.log(`⚠️ Warning: Could not find report type option: "${typeName}"`);
                    // [Debug] ถ้าหาไม่เจอ ให้ลองปริ้นท์รายการที่มีอยู่ทั้งหมดออกมาดูใน Log
                    /*
                    const allLabels = await page.evaluate(() => Array.from(document.querySelectorAll('label, span')).map(el => el.textContent.trim()));
                    console.log('Available options:', allLabels.slice(0, 20)); // ปริ้นท์มาดูบางส่วน
                    */
                }
                await new Promise(r => setTimeout(r, 500)); 
            } catch (e) {
                console.log(`⚠️ Error selecting ${typeName}:`, e.message);
            }
        }
        console.log('✅ Report Types Selected');

        // ---------------------------------------------------------
        // Step 3: Setting Date Range & Search
        // ---------------------------------------------------------
        console.log('3️⃣ Step 3: Setting Date Range 06:00 - 18:00...');
        const todayStr = getTodayFormatted();
        const startDateTime = `${todayStr} 06:00`;
        const endDateTime = `${todayStr} 18:00`;

        await page.evaluate(() => document.getElementById('date9').value = '');
        await page.type('#date9', startDateTime);

        await page.evaluate(() => document.getElementById('date10').value = '');
        await page.type('#date10', endDateTime);
        
        console.log('   Clicking Search to update report...');
        try {
            const searchBtnXPath = "//*[contains(text(), 'ค้นหา')] | //span[contains(@class, 'icon-search')]";
            const searchBtns = await page.$$(`xpath/${searchBtnXPath}`);
            
            if (searchBtns.length > 0) {
                await searchBtns[0].click();
            } else {
                await page.click('td:nth-of-type(5) > span');
            }
            console.log('   Waiting for report data to update...');
            await new Promise(r => setTimeout(r, 10000)); 
        } catch (e) {
            console.log('⚠️ Warning: Could not click Search button.', e.message);
        }

        // ---------------------------------------------------------
        // Step 4: Export Excel
        // ---------------------------------------------------------
        console.log('4️⃣ Step 4: Clicking Export/Excel...');
        
        cleanDownloadFolder(downloadPath);

        const excelBtnSelector = '#btnexport, button[title="Excel"], ::-p-aria(Excel)';
        await page.waitForSelector(excelBtnSelector, { visible: true, timeout: 60000 });
        
        await page.evaluate(() => {
            const btn = document.querySelector('#btnexport') || document.querySelector('button[title="Excel"]');
            if(btn) btn.click();
        });
        
        console.log('   Waiting for download (20s)...');
        await new Promise(r => setTimeout(r, 20000));

        // ---------------------------------------------------------
        // Step 5: Email & Cleanup
        // ---------------------------------------------------------
        console.log('5️⃣ Step 5: Processing email...');
        
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

            console.log('   Deleting downloaded file...');
            try {
                fs.unlinkSync(filePath);
                console.log('✅ File deleted successfully.');
            } catch (err) {
                console.error('⚠️ Error deleting file:', err);
            }

        } else {
            console.log('❌ No file downloaded to send.');
            throw new Error('Download failed or no file found');
        }

        console.log('🎉 Script completed successfully.');

    } catch (error) {
        console.error('❌ Error occurred:', error);
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
    console.log('📧 Email sent: ' + info.response);
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
