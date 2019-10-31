// 修改变量
const beginNumber = 1;
const endNumber = 365;

// 下面不要修改了
const Excel = require('exceljs');

console.log("hello");

var workbook = new Excel.Workbook();
var sheet = workbook.addWorksheet('plan');


sheet.addRow(["艾宾浩斯遗忘曲线复习计划表"]);
sheet.mergeCells('A1:F1');

sheet.addRow(["序号","学习日期","学习内容","时间","复习点","复习后打钩"]);

var rowNumber = 3;
for(var seq = beginNumber; seq <= endNumber; seq++) {
    var day1 = seq > 1 ? seq - 1 : "-";
    var day2 = seq > 2 ? seq - 2 : "-";
    var day4 = seq > 4 ? seq - 4 : "-";
    var day7 = seq > 7 ? seq - 7 : "-";
    var day15 = seq > 15 ? seq - 15 : "-";
    var month1 = seq > 30 ? seq - 30 : "-";
    var month2 = seq > 60 ? seq - 60 : "-";
    var month6 = seq > 180 ? seq - 180 : "-";

    sheet.addRow([seq, "月日", "", "12小时", seq]);
    sheet.addRow([seq, "月日", "", "1天", day1]);
    sheet.addRow([seq, "月日", "", "2天", day2]);
    sheet.addRow([seq, "月日", "", "4天", day4]);
    sheet.addRow([seq, "月日", "", "7天", day7]);
    sheet.addRow([seq, "月日", "", "15天", day15]);
    sheet.addRow([seq, "月日", "", "1月", month1]);
    sheet.addRow([seq, "月日", "", "3月", month2]);
    sheet.addRow([seq, "月日", "", "6月", month6]);

    sheet.mergeCells('A' + rowNumber + ':A' + (rowNumber + 8));
    sheet.mergeCells('B' + rowNumber + ':B' + (rowNumber + 8));
    sheet.mergeCells('C' + rowNumber + ':C' + (rowNumber + 8));

    rowNumber += 9;
}

// you can create xlsx file now.
workbook.xlsx.writeFile("/tmp/k8/ebbinghaus_plan.xlsx").then(function() {
    console.log("xls file is written.");
});