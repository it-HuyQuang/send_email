const express = require('express');
const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');
const pLimit = (require('p-limit').default || require('p-limit'));
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

// --- Progress tracking (SSE) ---
const jobs = new Map();

// --- Trang chính: gửi phiếu lương từ Excel ---
app.get('/', (req, res) => {
    res.render('bulk', { results: null, preview: null, error: null });
});

app.post('/preview', upload.single('excelFile'), async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(req.file.path);
        const employees = parseEmployees(workbook);

        const preview = employees.map(emp => ({
            ...emp,
            payslipHtml: buildPayslipHtml(emp, req.body.month, req.body.year),
        }));

        res.render('bulk', {
            results: null,
            preview,
            error: null,
            filePath: req.file.path,
            month: req.body.month,
            year: req.body.year,
        });
    } catch (error) {
        console.error(error);
        res.render('bulk', { results: null, preview: null, error: 'Lỗi đọc file: ' + error.message });
    }
});

// Gửi email - trả về jobId, xử lý background
app.post('/send', async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(req.body.filePath);
        const employees = parseEmployees(workbook);
        const { month, year } = req.body;

        const jobId = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
        const job = {
            total: employees.length,
            completed: 0,
            results: [],
            done: false,
        };
        jobs.set(jobId, job);

        // Process with p-limit concurrency
        const limit = pLimit(CONCURRENCY);
        const tasks = employees.map(emp => limit(async () => {
            const html = buildPayslipHtml(emp, month, year);
            try {
                await transporter.sendMail({
                    from: process.env.EMAIL_USER,
                    to: emp.gmail,
                    subject: `Phiếu lương tháng ${month}/${year} - ${emp.hoTen}`,
                    html,
                });
                job.results.push({ name: emp.hoTen, email: emp.gmail, status: 'success' });
            } catch (err) {
                job.results.push({ name: emp.hoTen, email: emp.gmail, status: 'fail', error: err.message });
            }
            job.completed++;
        }));

        Promise.all(tasks).then(() => {
            job.done = true;
            // Cleanup job sau 5 phút
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

// --- Helper: đọc Sheet 1 ---
function parseEmployees(workbook) {
    const sheet = workbook.getWorksheet(1);

    let headerRow = null;
    let headerMap = {};
    sheet.eachRow((row, rowNum) => {
        if (headerRow) return;
        row.eachCell((cell) => {
            const val = String(cell.value).trim().toUpperCase();
            if (val === 'STT') {
                headerRow = rowNum;
            }
        });
        if (headerRow && rowNum === headerRow) {
            row.eachCell((cell, colNum) => {
                headerMap[String(cell.value).trim().toUpperCase()] = colNum;
            });
        }
    });

    if (!headerRow) throw new Error('Không tìm thấy header row (cột STT) trong Sheet 1');

    const findCol = (exact, ...partials) => {
        for (const key of exact) {
            if (headerMap[key]) return headerMap[key];
        }
        for (const partial of partials) {
            const found = Object.keys(headerMap).find(k => k.includes(partial));
            if (found) return headerMap[found];
        }
        return null;
    };

    const col = {
        stt: headerMap['STT'],
        ten: findCol(['TÊN NHÂN VIÊN', 'TEN NHAN VIEN', 'HỌ TÊN']),
        gmail: findCol(['GMAIL', 'EMAIL']),
        chucVu: findCol(['CHỨC VỤ', 'CHUC VU']),
        ngayCong: findCol(['NGÀY CÔNG', 'NGAY CONG']),
        pc: findCol(['PC', 'PHỤ CẤP']),
        phat: findCol(['PHẠT', 'PHAT']),
        luong: findCol([], 'LƯƠNG', 'LUONG'),
    };

    const employees = [];
    sheet.eachRow((row, rowNum) => {
        if (rowNum <= headerRow) return;
        const stt = getCellValue(row, col.stt);
        const gmail = getCellValue(row, col.gmail);
        if (!stt || !gmail) return;

        const luong = toNumber(getCellValue(row, col.luong));
        const pc = toNumber(getCellValue(row, col.pc));
        const phat = toNumber(getCellValue(row, col.phat));
        const netIncome = luong + pc - phat;

        employees.push({
            maNV: String(stt).trim(),
            hoTen: String(getCellValue(row, col.ten) || '').trim(),
            gmail: String(gmail).trim(),
            chucVu: String(getCellValue(row, col.chucVu) || '').trim(),
            ngayCong: toNumber(getCellValue(row, col.ngayCong)),
            luong,
            pc,
            phat,
            netIncome,
        });
    });

    return employees;
}

function getCellValue(row, colNum) {
    if (!colNum) return null;
    const cell = row.getCell(colNum);
    const v = cell.value;
    if (v == null) return null;
    if (typeof v !== 'object') return v;

    // Formula cell: { formula, result }
    if (v.result !== undefined) return v.result;

    // Hyperlink cell: { text, hyperlink }
    if (v.text !== undefined) {
        const text = v.text;
        // RichText: { richText: [{ text: "value" }] }
        if (typeof text === 'object' && text.richText) {
            return text.richText.map(r => r.text).join('');
        }
        return String(text);
    }

    // RichText directly: { richText: [...] }
    if (v.richText) {
        return v.richText.map(r => r.text).join('');
    }

    return String(v);
}

function toNumber(val) {
    const n = Number(val);
    return isNaN(n) ? 0 : n;
}

function formatVND(num) {
    return Number(num).toLocaleString('vi-VN');
}

// --- Payslip HTML - Mobile responsive ---
function buildPayslipHtml(emp, month, year) {
    const mm = String(month || '--').padStart(2, '0');
    const yyyy = year || '--';
    return `
    <!DOCTYPE html>
    <html>
    <head><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>
    <body style="margin:0;padding:0;background:#f5f5f5;">
    <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;background:#f5f5f5;">
        <tr><td style="padding:16px 8px;">
            <table cellpadding="0" cellspacing="0" style="font-family:'Times New Roman',Georgia,serif;width:100%;max-width:600px;margin:0 auto;border-collapse:collapse;background:#ffffff;border:1px solid #999;">

                <!-- PAYSLIP Header -->
                <tr>
                    <td colspan="2" style="text-align:center;padding:16px 12px 4px;font-size:20px;font-weight:bold;letter-spacing:2px;">
                        PAYSLIP
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="text-align:center;padding:2px 12px 14px;font-style:italic;color:#666;font-size:14px;border-bottom:2px solid #333;">
                        Month: ${mm}, Year: ${yyyy}
                    </td>
                </tr>

                <!-- EMPLOYEE INFORMATION -->
                <tr>
                    <td colspan="2" style="background:#d9e1f2;padding:8px 12px;font-weight:bold;font-size:13px;text-align:center;border-bottom:1px solid #999;">
                        EMPLOYEE INFORMATION &ndash; TH&Ocirc;NG TIN NH&Acirc;N SỰ
                    </td>
                </tr>
                <tr>
                    <td style="padding:6px 12px;font-size:12px;color:#555;width:40%;">Full name / Họ v&agrave; t&ecirc;n:</td>
                    <td style="padding:6px 12px;font-size:13px;font-weight:bold;word-break:break-word;">${emp.hoTen}</td>
                </tr>
                <tr>
                    <td style="padding:6px 12px;font-size:12px;color:#555;">M&atilde; NV:</td>
                    <td style="padding:6px 12px;font-size:13px;font-weight:bold;color:#0066cc;">${emp.maNV}</td>
                </tr>
                <tr>
                    <td style="padding:6px 12px;font-size:12px;color:#555;">Position / Chức danh:</td>
                    <td style="padding:6px 12px;font-size:12px;">${emp.chucVu}</td>
                </tr>
                <tr>
                    <td style="padding:6px 12px;font-size:12px;color:#555;border-bottom:2px solid #333;">Bank account / T&agrave;i khoản ng&acirc;n h&agrave;ng:</td>
                    <td style="padding:6px 12px;border-bottom:2px solid #333;"></td>
                </tr>

                <!-- INCOME - THU NHẬP -->
                <tr>
                    <td colspan="2" style="background:#d9e1f2;padding:8px 12px;font-weight:bold;font-size:14px;text-align:center;border-bottom:1px solid #999;">
                        INCOME &ndash; THU NHẬP
                    </td>
                </tr>
                <tr>
                    <td style="padding:8px 12px;font-size:13px;border-bottom:1px solid #bbb;">
                        Ng&agrave;y l&agrave;m việc: <strong>${Number(emp.ngayCong).toFixed(2)}</strong>
                    </td>
                    <td style="padding:8px 12px;font-size:13px;text-align:right;border-bottom:1px solid #bbb;">
                        Th&agrave;nh tiền: <strong>${formatVND(emp.luong)}</strong>
                    </td>
                </tr>
                <tr>
                    <td style="padding:8px 12px;font-size:13px;border-bottom:1px solid #bbb;">
                        Phụ cấp tr&aacute;ch nhiệm
                    </td>
                    <td style="padding:8px 12px;font-size:13px;text-align:right;border-bottom:1px solid #bbb;">
                        ${emp.pc ? formatVND(emp.pc) : ''}
                    </td>
                </tr>
                <tr>
                    <td style="padding:8px 12px;font-size:13px;border-bottom:1px solid #bbb;">
                        Phụ cấp v&eacute; xe
                    </td>
                    <td style="padding:8px 12px;font-size:13px;text-align:right;border-bottom:1px solid #bbb;"></td>
                </tr>
                <tr>
                    <td style="padding:8px 12px;font-size:13px;border-bottom:2px solid #333;">
                        Phạt đi muộn
                    </td>
                    <td style="padding:8px 12px;font-size:13px;text-align:right;border-bottom:2px solid #333;color:#cc0000;">
                        ${emp.phat ? formatVND(emp.phat) + ' đ' : ''}
                    </td>
                </tr>

                <!-- NET INCOME -->
                <tr>
                    <td style="padding:12px;font-weight:bold;font-size:12px;border-bottom:2px solid #333;vertical-align:middle;">
                        NET INCOME &ndash; LƯƠNG THỰC LĨNH<br>
                        <span style="font-size:11px;color:#555;">(Currency / Đơn vị thanh to&aacute;n: VND)</span>
                    </td>
                    <td style="padding:12px;text-align:right;border-bottom:2px solid #333;vertical-align:middle;">
                        <div style="display:inline-block;font-size:22px;font-weight:bold;color:#cc0000;background:#fff2cc;padding:8px 16px;border:2px solid #cc0000;">
                            ${formatVND(emp.netIncome)}đ
                        </div>
                    </td>
                </tr>

                <!-- Ghi chú -->
                <tr>
                    <td colspan="2" style="padding:10px 12px 4px;font-size:11px;color:#cc0000;font-style:italic;">
                        * Ghi ch&uacute;:<br>
                        Th&ocirc;ng tin lương phải được bảo mật tuyệt đối. C&aacute; nh&acirc;n n&agrave;o v&ocirc; t&igrave;nh hoặc cố &yacute; l&agrave;m lộ th&ocirc;ng tin lương của c&aacute; nh&acirc;n hay đồng nghiệp sẽ bị xử l&yacute;.
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="padding:4px 12px;font-size:12px;">
                        <strong>Mọi thắc mắc vui l&ograve;ng li&ecirc;n hệ P. HCKT:</strong>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="padding:4px 12px;font-size:12px;font-weight:bold;">
                        Bạch Mai Ng&acirc;n - 0398 210 432
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="padding:4px 12px 14px;font-size:12px;font-style:italic;">
                        Thank you for your consideration. / Xin ch&acirc;n th&agrave;nh cảm ơn.
                    </td>
                </tr>

            </table>
        </td></tr>
    </table>
    </body>
    </html>`;
}

app.listen(PORT, () => {
    console.log(`Server đang chạy tại http://localhost:${PORT}`);
});
