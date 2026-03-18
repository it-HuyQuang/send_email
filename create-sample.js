const ExcelJS = require('exceljs');

async function createSample() {
    const workbook = new ExcelJS.Workbook();

    // Sheet 1: Bảng thanh toán tiền lương
    const sheet1 = workbook.addWorksheet('Bảng lương');

    // Row 1: Title
    sheet1.mergeCells('A1:H1');
    sheet1.getCell('A1').value = 'BẢNG THANH TOÁN TIỀN LƯƠNG';
    sheet1.getCell('A1').font = { bold: true, size: 14 };
    sheet1.getCell('A1').alignment = { horizontal: 'center' };

    // Row 2: Column numbers
    const colNums = ['1', '2', '3', '4', '5', '6', '7', '8'];
    const row2 = sheet1.getRow(2);
    colNums.forEach((n, i) => { row2.getCell(i + 1).value = n; });
    row2.alignment = { horizontal: 'center' };

    // Row 3: Headers
    const headers = ['STT', 'TÊN NHÂN VIÊN', 'GMAIL', 'CHỨC VỤ', 'NGÀY CÔNG', 'PC', 'PHẠT', 'LƯƠNG ĐỢT 1 T2'];
    const row3 = sheet1.getRow(3);
    headers.forEach((h, i) => {
        row3.getCell(i + 1).value = h;
    });
    row3.font = { bold: true, color: { argb: 'FFFF0000' } };
    row3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };

    // Column widths
    sheet1.getColumn(1).width = 10;
    sheet1.getColumn(2).width = 28;
    sheet1.getColumn(3).width = 30;
    sheet1.getColumn(4).width = 15;
    sheet1.getColumn(5).width = 14;
    sheet1.getColumn(6).width = 14;
    sheet1.getColumn(7).width = 12;
    sheet1.getColumn(8).width = 18;

    // Data rows
    const data = [
        ['NV005', 'TRỊNH THỊ NGUYỆT', 'chmaingan.yac@gmail.com', 'Phó Lead', 25, 1000000, 50000, 8840000],
        ['NV006', 'NGUYỄN THỊ THUỶ LINH', 'etoanmia.inc@gmail.com', 'Chính thức', 24, 0, 0, 6230000],
        ['NV007', 'NGUYỄN VĂN TRƯỜNG', 'etoanhdigroup@gmail.com', 'Phó Lead', 26, 1000000, 0, 7310000],
        ['NV008', 'NGUYỄN THỊ MINH NGỌC', 'bachmaingan@gmail.com', 'Phó Lead', 25, 1000000, 100000, 9080000],
    ];

    data.forEach(d => {
        sheet1.addRow(d);
    });

    // Format number columns
    for (let r = 4; r <= 7; r++) {
        sheet1.getRow(r).getCell(6).numFmt = '#,##0';
        sheet1.getRow(r).getCell(7).numFmt = '#,##0';
        sheet1.getRow(r).getCell(8).numFmt = '#,##0';
    }

    await workbook.xlsx.writeFile('sample.xlsx');
    console.log('Đã tạo file sample.xlsx thành công!');
}

createSample();
