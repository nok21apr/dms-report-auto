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
        headless: true, // ตั้งเป็น false เพื่อดูการทำงานตอนเทสได้
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--start-maximized'
        ]
    });
    
    const page = await browser.newPage();
    
    // --- Setup ---
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
        // Step 2.5: Select Truck "ทั้งหมด"
        // ---------------------------------------------------------
        console.log('   Selecting Truck "ทั้งหมด"...');
        await page.waitForSelector('#ddl_truck', { visible: true, timeout: 60000 });

        await page.waitForFunction(() => {
            const select = document.getElementById('ddl_truck');
            if (!select) return false;
            return Array.from(select.options).some(opt => opt.text.includes('ทั้งหมด') || opt.text.toLowerCase().includes('all'));
        }, { timeout: 60000 });

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
        // Step 2.6: Select Report Types (Force Click JS Method)
        // ---------------------------------------------------------
        console.log('   Selecting 3 Report Types...');
        
        // 1. พยายามเปิด Dropdown ก่อน (สำคัญมาก เพราะถ้าไม่เปิด Element อาจจะยังไม่ถูกสร้าง)
        try {
            console.log('      Attempting to open Status Dropdown...');
            // ลองคลิกที่ element ที่น่าจะเป็นตัวเปิด Dropdown สถานะ (เดาจากชื่อไฟล์ ddl_status หรือ label)
            await page.evaluate(() => {
                // ลองหา ddl_status, ddlStatus หรือ element ที่ใกล้กับคำว่า "สถานะ"
                const possibleIds = ['ddl_status', 'ddlStatus', 'status_list', 'div_status'];
                let opened = false;
                
                // ลองคลิก ID ที่น่าจะเป็นไปได้
                for(let id of possibleIds) {
                    const el = document.getElementById(id);
                    if(el) { el.click(); opened = true; break; }
                }
                
                // ถ้ายังไม่เจอ ลองหาจาก Label
                if(!opened) {
                    const labels = Array.from(document.querySelectorAll('label, span, div'));
                    const statusLabel = labels.find(l => l.innerText && (l.innerText.includes('สถานะ') || l.innerText.includes('ชนิดรายงาน')));
                    if(statusLabel) {
                        // คลิก element ถัดไป หรือ input ใกล้ๆ
                        const nextEl = statusLabel.nextElementSibling || statusLabel.parentElement.querySelector('div[class*="arrow"], div[class*="dropdown"]');
                        if(nextEl) nextEl.click();
                    }
                }
            });
            // รอให้ Animation ของ Dropdown ทำงานสักนิด
            await new Promise(r => setTimeout(r, 1000));
        } catch(e) {
            console.log('      ⚠️ Could not explicitly open dropdown (might be already open or non-standard).');
        }

        // 2. ใช้ JS ค้นหา Text และคลิกเลย (Force Click)
        const reportKeywords = ["ระดับ 1", "ระดับ 2", "หาวนอน"];
        
        for (const keyword of reportKeywords) {
            try {
                console.log(`      Searching for "${keyword}"...`);
                
                const found = await page.evaluate((kw) => {
                    // ใช้ XPath ค้นหา text node ที่มีคำนั้นอยู่
                    const xpath = `//*[contains(text(), '${kw}')]`;
                    const result = document.evaluate(xpath, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
                    
                    if (result.snapshotLength > 0) {
                        // วนลูปหา element ที่เหมาะสมที่สุดที่จะคลิก (เช่น div, span, li)
                        for (let i = 0; i < result.snapshotLength; i++) {
                            let el = result.snapshotItem(i);
                            
                            // เดินขึ้นไปหา Container ที่คลิกได้ (เช่นถ้า text อยู่ใน span เล็กๆ ให้คลิก div ที่หุ้มอยู่)
                            // เช็คว่าเป็น Dropdown Item หรือ Checkbox หรือไม่
                            while (el && el.tagName !== 'BODY') {
                                // ถ้าเจอ Checkbox
                                if (el.tagName === 'INPUT' && el.type === 'checkbox') {
                                    if(!el.checked) el.click();
                                    return true;
                                }
                                // ถ้าเจอ List Item (div/li)
                                if (el.tagName === 'LI' || (el.tagName === 'DIV' && (el.className.includes('item') || el.className.includes('list')))) {
                                    el.click();
                                    return true;
                                }
                                el = el.parentElement;
                            }
                            
                            // ถ้าไม่เจอ parent ที่ชัดเจน ให้คลิกที่ตัว element เองเลย
                            result.snapshotItem(i).click();
                            return true;
                        }
                    }
                    return false;
                }, keyword);

                if (found) {
                    console.log(`      ✅ Clicked option containing "${keyword}"`);
                } else {
                    console.log(`      ⚠️ Warning: Text "${keyword}" not found in DOM.`);
                }
                
                await new Promise(r => setTimeout(r, 500)); // เว้นจังหวะ
            } catch (e) {
                console.log(`      ❌ Error processing "${keyword}":`, e.message);
            }
        }
        console.log('✅ Report Types Selection Finished');

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
