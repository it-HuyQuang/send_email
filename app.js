const express = require('express');
const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const pLimit = (require('p-limit').default || require('p-limit'));
const puppeteer = require('puppeteer-core');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3003;
const CONCURRENCY = Number(process.env.EMAIL_CONCURRENCY) || 5;

app.set('view engine', 'ejs');
app.use(express.urlencoded({ extended: true }));

const upload = multer({
    dest: path.join(__dirname, 'uploads'),
    fileFilter(req, file, cb) {
        if (file.originalname.match(/\.xlsx$/)) {
            cb(null, true);
        } else {
            cb(new Error('Chỉ chấp nhận file .xlsx'));
        }
    },
});

const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
    },
});

// Logo VERA - đọc base64 để nhúng vào payslip
const logoPath = path.join(__dirname, 'image copy 3.png');
const logoBase64 = fs.existsSync(logoPath)
    ? fs.readFileSync(logoPath).toString('base64')
    : '';
const logoDataUri = logoBase64 ? `data:image/png;base64,${logoBase64}` : '';

// Tìm Chrome cho puppeteer PDF
function findChrome() {
    const paths = [
        process.env.CHROME_PATH,
        'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
        '/usr/bin/google-chrome',
        '/usr/bin/chromium-browser',
    ].filter(Boolean);
    for (const p of paths) {
        if (fs.existsSync(p)) return p;
    }
    return null;
}

let browserInstance;
async function getBrowser() {
    if (!browserInstance) {
        const chromePath = findChrome();
        if (!chromePath) throw new Error('Chrome không tìm thấy. Đặt CHROME_PATH trong .env');
        browserInstance = await puppeteer.launch({
            executablePath: chromePath,
            headless: 'new',
            args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu', '--window-position=-9999,-9999'],
        });
        browserInstance.on('disconnected', () => { browserInstance = null; });
    }
    return browserInstance;
}

async function htmlToPdf(html) {
    const browser = await getBrowser();
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'domcontentloaded', timeout: 60000 });
    const pdf = await page.pdf({
        format: 'A4',
        printBackground: true,
        margin: { top: '10mm', right: '10mm', bottom: '10mm', left: '10mm' },
    });
    await page.close();
    return pdf;
}

// --- Progress tracking (SSE) ---
const jobs = new Map();

// --- Routes ---
app.get('/', (req, res) => {
    res.render('bulk', { results: null, preview: null, error: null });
});

app.post('/preview', upload.single('excelFile'), async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(req.file.path);
        const dot = req.body.dot || '2';
        const month = req.body.month;
        const year = req.body.year;
        const deadlineDate = req.body.deadlineDate || '';
        const deadlineHour = req.body.deadlineHour || '';
        const deadlineMinute = req.body.deadlineMinute || '';

        const employees = dot === '2'
            ? parseEmployeesDot2(workbook)
            : parseEmployeesDot1(workbook);

        const preview = employees.map(emp => ({
            ...emp,
            payslipHtml: dot === '2'
                ? buildPayslipHtmlDot2(emp, month, year)
                : buildPayslipHtmlDot1(emp, month, year),
        }));

        res.render('bulk', {
            results: null,
            preview,
            error: null,
            filePath: req.file.path,
            month,
            year,
            dot,
            deadlineDate,
            deadlineHour,
            deadlineMinute,
        });
    } catch (error) {
        console.error(error);
        res.render('bulk', { results: null, preview: null, error: 'Lỗi đọc file: ' + error.message });
    }
});

app.post('/send', async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(req.body.filePath);
        const { month, year, dot, deadlineDate, deadlineHour, deadlineMinute } = req.body;

        const employees = dot === '2'
            ? parseEmployeesDot2(workbook)
            : parseEmployeesDot1(workbook);

        const jobId = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
        const job = { total: employees.length, completed: 0, results: [], done: false };
        jobs.set(jobId, job);

        const limit = pLimit(CONCURRENCY);
        const mm = String(month).padStart(2, '0');

        // Format deadline string
        let deadlineStr = '';
        if (deadlineDate && deadlineHour) {
            const [y, m, d] = deadlineDate.split('-');
            deadlineStr = `${deadlineHour}h${deadlineMinute || '00'} ngày ${d}/${m}/${y}`;
        }

        const tasks = employees.map(emp => limit(async () => {
            try {
                const payslipHtml = dot === '2'
                    ? buildPayslipHtmlDot2(emp, month, year)
                    : buildPayslipHtmlDot1(emp, month, year);

                // Convert HTML → PDF
                const pdfBuffer = await htmlToPdf(payslipHtml);

                // Email body
                const emailBody = buildEmailBody(dot, mm, year, deadlineStr);
                const subject = `PHIẾU LƯƠNG ĐỢT ${dot} THÁNG ${mm}/${year} - ${emp.hoTen}`;

                // Tên file PDF
                const safeName = removeVietnameseTones(emp.hoTen).replace(/\s+/g, '-');
                const filename = `Phieu-luong-dot-${dot}-${mm}-${year}-${safeName}.pdf`;

                await transporter.sendMail({
                    from: process.env.EMAIL_USER,
                    to: emp.gmail,
                    subject,
                    html: emailBody,
                    attachments: [{
                        filename,
                        content: pdfBuffer,
                        contentType: 'application/pdf',
                    }],
                });
                job.results.push({ name: emp.hoTen, email: emp.gmail, status: 'success' });
            } catch (err) {
                job.results.push({ name: emp.hoTen, email: emp.gmail, status: 'fail', error: err.message });
            }
            job.completed++;
        }));

        Promise.all(tasks).then(() => {
            job.done = true;
            setTimeout(() => jobs.delete(jobId), 5 * 60 * 1000);
        });

        res.json({ jobId });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Lỗi gửi: ' + error.message });
    }
});

// SSE progress endpoint
app.get('/progress/:jobId', (req, res) => {
    const jobId = req.params.jobId;
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders();

    const interval = setInterval(() => {
        const job = jobs.get(jobId);
        if (!job) {
            res.write('data: {"error":"Job not found"}\n\n');
            clearInterval(interval);
            res.end();
            return;
        }
        res.write(`data: ${JSON.stringify({
            total: job.total,
            completed: job.completed,
            done: job.done,
            results: job.done ? job.results : null,
        })}\n\n`);
        if (job.done) {
            clearInterval(interval);
            res.end();
        }
    }, 500);
    req.on('close', () => clearInterval(interval));
});

// ===================== HELPERS =====================

function getCellValue(row, colNum) {
    if (!colNum) return null;
    const cell = row.getCell(colNum);
    const v = cell.value;
    if (v == null) return null;
    if (typeof v !== 'object') return v;
    if (v.result !== undefined) return v.result;
    if (v.text !== undefined) {
        const text = v.text;
        if (typeof text === 'object' && text.richText) {
            return text.richText.map(r => r.text).join('');
        }
        return String(text);
    }
    if (v.richText) return v.richText.map(r => r.text).join('');
    return String(v);
}

function toNumber(val) {
    const n = Number(val);
    return isNaN(n) ? 0 : Math.round(n);
}

function formatVND(num) {
    return Math.round(Number(num)).toLocaleString('vi-VN');
}

function removeVietnameseTones(str) {
    return str.normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd').replace(/Đ/g, 'D');
}

// ===================== PARSE HEADERS =====================

function parseCellText(cell) {
    const v = cell.value;
    if (v == null) return '';
    if (typeof v !== 'object') return String(v).trim();
    if (v.result !== undefined) return String(v.result).trim();
    if (v.richText) return v.richText.map(r => r.text).join('').trim();
    if (v.text !== undefined) {
        const text = v.text;
        if (typeof text === 'object' && text.richText) {
            return text.richText.map(r => r.text).join('').trim();
        }
        return String(text).trim();
    }
    return String(v).trim();
}

function parseHeaders(workbook) {
    let targetSheet = null;
    let headerRow = null;
    const headers = {};

    for (const sheet of workbook.worksheets) {
        if (headerRow) break;
        sheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
            if (headerRow) return;
            row.eachCell({ includeEmpty: false }, (cell, colNum) => {
                if (headerRow) return;
                const val = parseCellText(cell).toUpperCase();
                if (val === 'STT') {
                    headerRow = rowNum;
                    targetSheet = sheet;
                }
            });
            if (headerRow && rowNum === headerRow) {
                row.eachCell({ includeEmpty: false }, (cell, colNum) => {
                    const val = parseCellText(cell).toUpperCase();
                    if (val) headers[val] = colNum;
                });
            }
        });
    }

    if (!headerRow || !targetSheet) throw new Error('Không tìm thấy header row (cột STT) trong file Excel');
    console.log('Sheet:', targetSheet.name, '| Header row:', headerRow, '| Columns:', headers);
    return { sheet: targetSheet, headerRow, headers };
}

function findCol(headers, exactNames, partialName, excludePatterns) {
    for (const name of exactNames) {
        if (headers[name] !== undefined) return headers[name];
    }
    if (partialName) {
        const ex = excludePatterns || [];
        const key = Object.keys(headers).find(k =>
            k.includes(partialName) && !ex.some(e => k.includes(e))
        );
        if (key) return headers[key];
    }
    return null;
}

// ===================== ĐỢT 1 =====================

function parseEmployeesDot1(workbook) {
    const { sheet, headerRow, headers } = parseHeaders(workbook);

    const col = {
        stt: headers['STT'],
        ten: findCol(headers, ['TÊN NHÂN VIÊN', 'TEN NHAN VIEN', 'HỌ TÊN', 'HO TEN']),
        gmail: findCol(headers, ['GMAIL', 'EMAIL']),
        chucVu: findCol(headers, ['CHỨC VỤ', 'CHUC VU']),
        ngayCong: findCol(headers, ['NGÀY CÔNG', 'NGAY CONG']),
        lcDot1: findCol(headers, ['LC ĐỢT 1', 'LC DOT 1'], 'LC ĐỢT'),
        phuCapTN: findCol(headers, ['PHỤ CẤP TN', 'PHU CAP TN']),
        phuCapXe: findCol(headers, ['PHỤ CẤP XE', 'PHU CAP XE']),
        phatDiMuon: findCol(headers, ['PHẠT ĐI MUỘN', 'PHAT DI MUON'], 'PHẠT ĐI MUỘN'),
    };

    const employees = [];
    sheet.eachRow((row, rowNum) => {
        if (rowNum <= headerRow) return;
        const stt = getCellValue(row, col.stt);
        const gmail = getCellValue(row, col.gmail);
        if (!stt || !gmail) return;
        if (String(stt).toUpperCase() === 'TỔNG') return;

        const lcDot1 = toNumber(getCellValue(row, col.lcDot1));
        const phuCapTN = toNumber(getCellValue(row, col.phuCapTN));
        const phuCapXe = toNumber(getCellValue(row, col.phuCapXe));
        const phatDiMuon = toNumber(getCellValue(row, col.phatDiMuon));

        employees.push({
            maNV: String(stt).trim(),
            hoTen: String(getCellValue(row, col.ten) || '').trim(),
            gmail: String(gmail).trim(),
            chucVu: String(getCellValue(row, col.chucVu) || '').trim(),
            ngayCong: toNumber(getCellValue(row, col.ngayCong)),
            lcDot1, phuCapTN, phuCapXe, phatDiMuon,
            netIncome: lcDot1 + phuCapTN + phuCapXe - phatDiMuon,
        });
    });
    return employees;
}

function buildPayslipHtmlDot1(emp, month, year) {
    const mm = String(month || '--').padStart(2, '0');
    const yyyy = year || '--';

    const fv = (val) => val ? formatVND(val) + ' đ' : '';

    return `<!DOCTYPE html>
<html>
<head><meta name="viewport" content="width=device-width,initial-scale=1.0">
<style>
    body { margin:0; padding:0; background:#fff; }
    table.payslip {
        width:100%; max-width:600px; margin:0 auto;
        font-family:'Times New Roman',Georgia,serif;
        border-collapse:collapse; background:#fff;
    }
    table.payslip td { vertical-align:top; }
    .bl { border-left:1px solid #bbb; }
    .br { border-right:1px solid #bbb; }
    .blr { border-left:1px solid #bbb; border-right:1px solid #bbb; }
    .hdr { background:#fff2cc; padding:8px 14px; font-weight:bold; text-align:center; border-top:1px solid #bbb; border-bottom:1px solid #bbb; border-left:1px solid #bbb; border-right:1px solid #bbb; }
    .info-label { padding:8px 14px; font-size:13px; color:#555; width:48%; border-bottom:1px solid #ddd; }
    .info-value { padding:8px 14px; font-size:13px; border-bottom:1px solid #ddd; }
    .item-label { padding:7px 14px; font-size:13px; border-bottom:1px solid #ddd; }
    .item-value { padding:7px 14px; font-size:13px; text-align:right; font-weight:bold; border-bottom:1px solid #ddd; }
</style>
</head>
<body>
<table class="payslip" cellpadding="0" cellspacing="0">

    <!-- HEADER: Logo + Company -->
    <tr>
        <td colspan="2" style="padding:16px 14px;border-left:1px solid #bbb;border-right:1px solid #bbb;border-top:1px solid #bbb;border-bottom:1px solid #bbb;">
            <div style="display:flex;align-items:center;gap:10px;">
                ${logoDataUri ? `<img src="${logoDataUri}" alt="VERA" style="height:70px;width:auto;">` : ''}
                <div>
                    <div style="font-size:20px;font-weight:bold;color:#b8860b;letter-spacing:1px;">VERA GROUP</div>
                    <div style="font-size:12px;color:#b8860b;margin-top:4px;">39 Lê Văn Lương - Thanh Xuân - Hà Nội</div>
                </div>
            </div>
        </td>
    </tr>

    <!-- PAYSLIP Title -->
    <tr>
        <td colspan="2" class="blr" style="text-align:center;padding:20px 14px 4px;font-size:22px;font-weight:bold;letter-spacing:2px;">
            PAYSLIP
        </td>
    </tr>
    <tr>
        <td colspan="2" class="blr" style="text-align:center;padding:2px 14px 16px;font-style:italic;color:#333;font-size:14px;border-bottom:2px solid #333;">
            Month: ${mm}, Year: ${yyyy}
        </td>
    </tr>

    <!-- EMPLOYEE INFORMATION -->
    <tr><td colspan="2" class="hdr" style="font-size:13px;border-top:2px solid #333;">EMPLOYEE INFORMATION – THÔNG TIN NHÂN SỰ</td></tr>
    <tr>
        <td class="info-label bl">Full name / Họ và tên:</td>
        <td class="info-value br" style="font-weight:bold;font-size:14px;">${emp.hoTen}</td>
    </tr>
    <tr>
        <td class="info-label bl">Mã NV:</td>
        <td class="info-value br" style="font-weight:bold;color:#cc0000;">${emp.maNV}</td>
    </tr>
    <tr>
        <td class="info-label bl">Position / Chức danh:</td>
        <td class="info-value br">${emp.chucVu}</td>
    </tr>
    <tr>
        <td class="info-label bl">Bank account / Tài khoản ngân hàng:</td>
        <td class="info-value br"></td>
    </tr>

    <!-- INCOME -->
    <tr><td colspan="2" class="hdr" style="font-size:14px;border-top:2px solid #333;">INCOME – THU NHẬP</td></tr>
    <tr>
        <td class="item-label bl">Ngày làm việc: <strong>${Number(emp.ngayCong).toFixed(2)}</strong></td>
        <td class="item-value br">Thành tiền: <strong>${formatVND(emp.lcDot1)} đ</strong></td>
    </tr>
    <tr><td class="item-label bl">Phụ cấp TN</td><td class="item-value br">${fv(emp.phuCapTN)}</td></tr>
    <tr><td class="item-label bl">Phụ cấp xe</td><td class="item-value br">${fv(emp.phuCapXe)}</td></tr>
    <tr>
        <td class="item-label bl" style="color:#cc0000;">Phạt đi muộn</td>
        <td class="item-value br" style="color:#cc0000;">${fv(emp.phatDiMuon)}</td>
    </tr>

    <!-- NET INCOME -->
    <tr>
        <td class="bl" style="padding:14px;font-weight:bold;font-size:13px;border-top:2px solid #333;border-bottom:2px solid #333;vertical-align:middle;">
            NET INCOME – LƯƠNG THỰC LĨNH<br>
            <span style="font-size:11px;color:#555;font-style:italic;">(Currency / Đơn vị thanh toán: VND)</span>
        </td>
        <td class="br" style="padding:14px;text-align:right;border-top:2px solid #333;border-bottom:2px solid #333;vertical-align:middle;">
            <div style="display:inline-block;font-size:24px;font-weight:bold;color:#cc0000;background:#fff2cc;padding:8px 18px;border:2px solid #cc0000;">
                ${formatVND(emp.netIncome)}đ
            </div>
        </td>
    </tr>

    <!-- Ghi chú -->
    <tr>
        <td colspan="2" class="blr" style="padding:12px 14px 4px;font-size:11px;color:#cc0000;font-style:italic;">
            * Ghi chú:<br>
            Thông tin lương phải được bảo mật tuyệt đối. Cá nhân nào vô tình hoặc cố ý làm lộ thông tin lương của cá nhân hay đồng nghiệp sẽ bị xử lý kỷ luật theo quy định.
        </td>
    </tr>
    <tr>
        <td colspan="2" class="blr" style="padding:4px 14px;font-size:12px;font-weight:bold;">
            Mọi thắc mắc vui lòng liên hệ P. HCKT:
        </td>
    </tr>
    <tr>
        <td colspan="2" class="blr" style="padding:4px 14px;font-size:12px;font-weight:bold;">
            Bạch Mai Ngân - 0398 210 432
        </td>
    </tr>
    <tr>
        <td colspan="2" class="blr" style="padding:4px 14px 16px;font-size:12px;font-style:italic;border-bottom:1px solid #bbb;">
            Thank you for your consideration. / Xin chân thành cảm ơn.
        </td>
    </tr>

</table>
</body>
</html>`;
}

// ===================== ĐỢT 2 =====================

function parseEmployeesDot2(workbook) {
    const { sheet, headerRow, headers } = parseHeaders(workbook);

    const col = {
        stt: headers['STT'],
        ten: findCol(headers, ['TÊN NHÂN VIÊN', 'TEN NHAN VIEN', 'HỌ TÊN']),
        gmail: findCol(headers, ['GMAIL', 'EMAIL']),
        chucVu: findCol(headers, ['CHỨC VỤ', 'CHUC VU']),
        doanhSo: findCol(headers, ['DOANH SỐ'], 'DOANH SỐ', ['TEAM']),
        doanhSoTeam: findCol(headers, ['DOANH SỐ TEAM']),
        cpqc: findCol(headers, ['CPQC'], 'CPQC', ['TEAM']),
        cpqcTeam: findCol(headers, ['CPQC TEAM']),
        lcDot2: findCol(headers, ['LC ĐỢT 2', 'LC DOT 2'], 'LC ĐỢT'),
        com: findCol(headers, ['COM'], 'COM', ['LEAD', 'TEAM', 'MISSION', 'GROSS']),
        comLead: findCol(headers, ['COM LEAD', 'COM TEAM']),
        thuongDS: findCol(headers, ['THƯỞNG DS', 'THUONG DS'], 'THƯỞNG DS'),
        thuongBestTop: findCol(headers, ['THƯỞNG BEST TOP', 'THƯỞNG BEST/TOP'], 'BEST'),
        thuongToiUu: findCol(headers, ['THƯỞNG TỐI ƯU'], 'TỐI ƯU'),
        thuongKhac: findCol(headers, ['THƯỞNG KHÁC'], 'THƯỞNG KHÁC'),
        phat: findCol(headers, ['PHẠT'], 'PHẠT', ['MUỘN', 'ĐI']),
        tongThuong: findCol(headers, ['TỔNG THƯỞNG'], 'TỔNG THƯỞNG'),
        phatDiMuon: findCol(headers, ['PHẠT ĐI MUỘN'], 'ĐI MUỘN'),
        luongDot2: findCol(headers, [], 'LƯƠNG ĐỢT 2'),
    };

    const employees = [];
    sheet.eachRow((row, rowNum) => {
        if (rowNum <= headerRow) return;
        const stt = getCellValue(row, col.stt);
        const gmail = getCellValue(row, col.gmail);
        if (!stt || !gmail) return;
        if (String(stt).toUpperCase() === 'TỔNG') return;

        const emp = {
            maNV: String(stt).trim(),
            hoTen: String(getCellValue(row, col.ten) || '').trim(),
            gmail: String(gmail).trim(),
            chucVu: String(getCellValue(row, col.chucVu) || '').trim(),
            doanhSo: toNumber(getCellValue(row, col.doanhSo)),
            doanhSoTeam: toNumber(getCellValue(row, col.doanhSoTeam)),
            cpqc: toNumber(getCellValue(row, col.cpqc)),
            cpqcTeam: toNumber(getCellValue(row, col.cpqcTeam)),
            lcDot2: toNumber(getCellValue(row, col.lcDot2)),
            com: toNumber(getCellValue(row, col.com)),
            comLead: toNumber(getCellValue(row, col.comLead)),
            thuongDS: toNumber(getCellValue(row, col.thuongDS)),
            thuongBestTop: toNumber(getCellValue(row, col.thuongBestTop)),
            thuongToiUu: toNumber(getCellValue(row, col.thuongToiUu)),
            thuongKhac: toNumber(getCellValue(row, col.thuongKhac)),
            phat: toNumber(getCellValue(row, col.phat)),
            tongThuong: toNumber(getCellValue(row, col.tongThuong)),
            phatDiMuon: toNumber(getCellValue(row, col.phatDiMuon)),
            luongDot2: toNumber(getCellValue(row, col.luongDot2)),
        };

        // Nếu không có cột LƯƠNG ĐỢT 2, tự tính
        if (!emp.luongDot2) {
            emp.luongDot2 = emp.lcDot2 + emp.com + emp.comLead + emp.thuongDS
                + emp.thuongBestTop + emp.thuongToiUu + emp.thuongKhac - emp.phat;
        }

        employees.push(emp);
    });
    return employees;
}

function buildPayslipHtmlDot2(emp, month, year) {
    const mm = String(month || '--').padStart(2, '0');
    const yyyy = year || '--';

    const fv = (val) => val ? formatVND(val) + ' đ' : '';
    const fvz = (val) => formatVND(val || 0) + ' đ';

    return `<!DOCTYPE html>
<html>
<head><meta name="viewport" content="width=device-width,initial-scale=1.0">
<style>
    body { margin:0; padding:0; background:#fff; }
    table.payslip {
        width:100%; max-width:600px; margin:0 auto;
        font-family:'Times New Roman',Georgia,serif;
        border-collapse:collapse; background:#fff;
    }
    table.payslip td { vertical-align:top; }
    .bl { border-left:1px solid #bbb; }
    .br { border-right:1px solid #bbb; }
    .blr { border-left:1px solid #bbb; border-right:1px solid #bbb; }
    .hdr { background:#fff2cc; padding:8px 14px; font-weight:bold; text-align:center; border-top:1px solid #bbb; border-bottom:1px solid #bbb; border-left:1px solid #bbb; border-right:1px solid #bbb; }
    .info-label { padding:8px 14px; font-size:13px; color:#555; width:48%; border-bottom:1px solid #ddd; }
    .info-value { padding:8px 14px; font-size:13px; border-bottom:1px solid #ddd; }
    .item-label { padding:7px 14px; font-size:13px; border-bottom:1px solid #ddd; }
    .item-value { padding:7px 14px; font-size:13px; text-align:right; font-weight:bold; border-bottom:1px solid #ddd; }
</style>
</head>
<body>
<table class="payslip" cellpadding="0" cellspacing="0">

    <!-- HEADER: Logo + Company -->
    <tr>
        <td colspan="2" style="padding:16px 14px;border-left:1px solid #bbb;border-right:1px solid #bbb;border-top:1px solid #bbb;border-bottom:1px solid #bbb;">
            <div style="display:flex;align-items:center;gap:10px;">
                ${logoDataUri ? `<img src="${logoDataUri}" alt="VERA" style="height:70px;width:auto;">` : ''}
                <div>
                    <div style="font-size:20px;font-weight:bold;color:#b8860b;letter-spacing:1px;">VERA GROUP</div>
                    <div style="font-size:12px;color:#b8860b;margin-top:4px;">39 Lê Văn Lương - Thanh Xuân - Hà Nội</div>
                </div>
            </div>
        </td>
    </tr>

    <!-- PAYSLIP Title (bordered box, golden bottom) -->
    <tr>
        <td colspan="2" class="blr" style="text-align:center;padding:20px 14px 4px;font-size:22px;font-weight:bold;letter-spacing:2px;">
            PAYSLIP
        </td>
    </tr>
    <tr>
        <td colspan="2" class="blr" style="text-align:center;padding:2px 14px 16px;font-style:italic;color:#333;font-size:14px;border-bottom:2px solid #333;">
            Month: ${mm}, Year: ${yyyy}
        </td>
    </tr>

    <!-- EMPLOYEE INFORMATION -->
    <tr><td colspan="2" class="hdr" style="font-size:13px;border-top:2px solid #333;">EMPLOYEE INFORMATION – THÔNG TIN NHÂN SỰ</td></tr>
    <tr>
        <td class="info-label bl">Full name / Họ và tên:</td>
        <td class="info-value br" style="font-weight:bold;font-size:14px;">${emp.hoTen}</td>
    </tr>
    <tr>
        <td class="info-label bl">Mã NV:</td>
        <td class="info-value br" style="font-weight:bold;color:#cc0000;">${emp.maNV}</td>
    </tr>
    <tr>
        <td class="info-label bl">Position / Chức danh:</td>
        <td class="info-value br">${emp.chucVu}</td>
    </tr>
    <tr>
        <td class="info-label bl">Bank account / Tài khoản ngân hàng:</td>
        <td class="info-value br"></td>
    </tr>

    <!-- Thưởng / Commission -->
    <tr><td colspan="2" class="hdr" style="font-size:14px;border-top:2px solid #333;">Thưởng / Commission</td></tr>
    <tr>
        <td colspan="2" class="blr" style="padding:10px 14px 6px;font-size:13px;font-weight:bold;border-bottom:1px solid #ddd;">
            A. Lương cứng đợt 2
        </td>
    </tr>

    <!-- Commission items -->
    <tr><td class="item-label bl">1. COM cá nhân</td><td class="item-value br">${fv(emp.com)}</td></tr>
    <tr><td class="item-label bl">2. COM Team</td><td class="item-value br">${fv(emp.comLead)}</td></tr>
    <tr><td class="item-label bl">3. Thưởng doanh số</td><td class="item-value br">${fv(emp.thuongDS)}</td></tr>
    <tr><td class="item-label bl">5. Thưởng BEST/TOP</td><td class="item-value br">${fv(emp.thuongBestTop)}</td></tr>
    <tr><td class="item-label bl">6. Thưởng tối ưu</td><td class="item-value br">${fv(emp.thuongToiUu)}</td></tr>
    <tr><td class="item-label bl">7. Thưởng khác</td><td class="item-value br">${fv(emp.thuongKhac)}</td></tr>
    <tr>
        <td class="item-label bl" style="color:#cc0000;">8. Phạt</td>
        <td class="item-value br" style="color:#cc0000;">${fv(emp.phat)}</td>
    </tr>
    <tr><td class="item-label bl">9. Tạm ứng</td><td class="item-value br"></td></tr>

    <!-- Doanh số info -->
    <tr>
        <td class="item-label bl" style="font-style:italic;">Doanh số:</td>
        <td class="item-value br" style="font-style:italic;font-weight:normal;">${fvz(emp.doanhSo)}</td>
    </tr>
    <tr>
        <td class="item-label bl" style="font-style:italic;">Doanh số team:</td>
        <td class="item-value br" style="font-style:italic;font-weight:normal;">${fvz(emp.doanhSoTeam)}</td>
    </tr>
    <tr>
        <td class="item-label bl" style="font-style:italic;">Chi phí quảng cáo:</td>
        <td class="item-value br" style="font-style:italic;font-weight:normal;">${fvz(emp.cpqc)}</td>
    </tr>
    <tr>
        <td class="item-label bl" style="font-style:italic;">Chi phí QC team:</td>
        <td class="item-value br" style="font-style:italic;font-weight:normal;">${fvz(emp.cpqcTeam)}</td>
    </tr>

    <!-- GROSS COMMISSION – TỔNG LƯƠNG -->
    <tr>
        <td class="bl" style="padding:14px;font-weight:bold;font-size:13px;border-top:2px solid #333;border-bottom:2px solid #333;vertical-align:middle;">
            GROSS COMMISSION – TỔNG LƯƠNG<br>
            <span style="font-size:11px;color:#555;font-style:italic;">(Currency / Đơn vị thanh toán: VND)</span>
        </td>
        <td class="br" style="padding:14px;text-align:right;border-top:2px solid #333;border-bottom:2px solid #333;vertical-align:middle;">
            <div style="display:inline-block;font-size:24px;font-weight:bold;color:#cc0000;background:#fff2cc;padding:8px 18px;border:2px solid #cc0000;">
                ${formatVND(emp.luongDot2)}đ
            </div>
        </td>
    </tr>

    <!-- Ghi chú -->
    <tr>
        <td colspan="2" class="blr" style="padding:12px 14px 4px;font-size:11px;color:#cc0000;font-style:italic;">
            * Ghi chú:<br>
            Thông tin lương phải được bảo mật tuyệt đối. Cá nhân nào vô tình hoặc cố ý làm lộ thông tin lương của cá nhân hay đồng nghiệp sẽ bị xử lý kỷ luật theo quy định.
        </td>
    </tr>
    <tr>
        <td colspan="2" class="blr" style="padding:4px 14px;font-size:12px;font-weight:bold;">
            Mọi thắc mắc vui lòng liên hệ P. HCKT:
        </td>
    </tr>
    <tr>
        <td colspan="2" class="blr" style="padding:4px 14px;font-size:12px;font-weight:bold;">
            Bạch Mai Ngân - 0398 210 432
        </td>
    </tr>
    <tr>
        <td colspan="2" class="blr" style="padding:4px 14px 16px;font-size:12px;font-style:italic;border-bottom:1px solid #bbb;">
            Thank you for your consideration. / Xin chân thành cảm ơn.
        </td>
    </tr>

</table>
</body>
</html>`;
}

// ===================== EMAIL BODY =====================

function buildEmailBody(dot, mm, year, deadlineStr) {
    return `<div style="font-family:'Segoe UI',Arial,Helvetica,sans-serif;font-size:15px;line-height:1.8;color:#333;">
    <p>Phòng HC - KT VERA GROUP gửi phiếu lương <b style="font-weight:700;">ĐỢT ${dot}</b> tháng <b style="font-weight:700;">${mm}/${year}</b></p>
    ${deadlineStr ? `<p>Mọi sự thắc mắc, khiếu nại liên quan đến các khoản thể hiện trong payslip, VERA-ers vui lòng liên hệ lại phòng HC - KT trước <b style="font-weight:700;"><i>${deadlineStr}</i></b> để được giải quyết.</p>` : ''}
    <p>Trân trọng,</p>
    <p>Phòng HC - KT.</p>
</div>`;
}

// ===================== START =====================

app.listen(PORT, () => {
    console.log(`Server đang chạy tại http://localhost:${PORT}`);
});
