// Function to get month name from month number
function getMonthName(month) {
    var monthNames = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"];
    return monthNames[month - 1];
}

function createMonthlySheet(year, month) {
    if (typeof year === 'undefined' || typeof month === 'undefined') {
        Logger.log('Exiting function early due to undefined arguments.');
        return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    sheet.getRange('A1:Z1000').clearContent();
    sheet.getRange("A1:H1").merge();

    var weekdays = ["שבת", "שישי", "חמישי", "רביעי", "שלישי", "שני", "ראשון"];
    sheet.getRange("B2:H2").setValues([weekdays]).setFontSize(14).setBackground("#E0E0E0");

    var daysInMonth = new Date(year, month, 0).getDate();
    var firstDayOfWeek = new Date(year, month - 1, 1).getDay();
    var date = 1;
    var lastRow;

    for (var row = 3; date <= daysInMonth; row += 3) {
        for (var col = 8; col >= 2; col--) {
            var cell = sheet.getRange(row, col);
            var curColDayOfWeek = 7 - (col - 1);

            if (firstDayOfWeek == curColDayOfWeek && date == 1) {
                cell.setValue(date++).setHorizontalAlignment('center');
                firstDayOfWeek = -1; // Reset so we don't hit this if statement again
            } else if (date > 1 && date <= daysInMonth) {
                cell.setValue(date++).setHorizontalAlignment('center');
            } else {
                cell.setBackground("#CCCCCC");
            }

            cell.setFontSize(12);
            var color;
            if (month >= 6 && month <= 8) {
                color = "#FFCC00";
            } else if (month >= 3 && month <= 5) {
                color = "#99FF99";
            } else if (month >= 9 && month <= 11) {
                color = "#FF9966";
            } else {
                color = "#99CCFF";
            }
            cell.setBackground(color);

            sheet.getRange(row + 1, col).setBackground("#FFFFFF").setFontSize(8);
            sheet.getRange(row + 2, col).setBackground("#e8dddc").setFontSize(8);
            sheet.setRowHeight(row, 20);
            sheet.setRowHeight(row + 1, 80);
            sheet.setRowHeight(row + 2, 20);
        }

        var sumFormulaForThisRow = `=SUM(B${row + 2}:H${row + 2})`;
        sheet.getRange(row + 2, 9).setFormula(sumFormulaForThisRow);
        lastRow = row + 2;
    }

    sheet.getRange(2, 2, lastRow - 1, 7).setBorder(true, true, true, true, false, true);

    sheet.deleteColumns(10, sheet.getMaxColumns() - 9);
    sheet.deleteRows(lastRow + 1, sheet.getMaxRows() - lastRow);

    var sumCell = sheet.getRange(lastRow + 1, 8);
    var formula = `=SUM(I5:I${lastRow})`;
    sumCell.setFormula(formula).setFontSize(12).setBackground("#E0E0E0");

    sheet.getRange(2, 9, lastRow).setBackground("#E0E0E0");
    sheet.getRange(lastRow + 1, 2, 1, 7).setBackground("#E0E0E0");

    sheet.deleteColumn(1);

    var monthName = getMonthName(month);
    sheet.getRange("A1").setValue(monthName).setFontWeight("bold").setFontSize(14).setHorizontalAlignment('center');
}

createMonthlySheet(2023, 11);
