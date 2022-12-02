const express = require('express');
const ExcelJS = require('exceljs');

const app = new express();

app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});

app.get('/download', (req, res) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('My Sheet');

    // add column headers
    worksheet.columns = [
        { header: 'Date', key: 'date' },
    ]

    // // add rows
    worksheet.addRow({ date: new Date() });
    // row.commit();

    // set cell value
    // worksheet.getCell(1, 1).value = new Date();

    // write to file
    // workbook.xlsx.writeFile('test.xlsx')

    // download without writing to file
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=' + 'test.xlsx');
    workbook.xlsx.write(res).then(() => {
        res.end();
    });
});

app.listen(2000, function () {
    console.log("Node Server is Running at Port 2000");
});