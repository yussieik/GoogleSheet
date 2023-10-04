function getMonthName(month) {
    const monthNames = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"];
    return monthNames[month - 1];
}

function setWeekdaysHeader(sheet) {
    const weekdays = ["שבת", "שישי", "חמישי", "רביעי", "שלישי", "שני", "ראשון"];
    sheet.getRange("B2:H2").setValues([weekdays]).setFontSize(14).setBackground("#E0E0E0");
}

function setMonthColor(month, cell) {
    let color;
    if (month >= 6 && month <= 8) color = "#FFCC00";
    else if (month >= 3 && month <= 5) color = "#99FF99";
    else if (month >= 9 && month <= 11) color = "#FF9966";
    else color = "#99CCFF";
    cell.setBackground(color);
}

function clearSheet(sheet) {
    sheet.getRange('A1:Z1000').clearContent();
    sheet.getRange("A1:H1").merge();
}

function createMonthlySheet(year, month) {

    var ss = SpreadsheetApp.create(getMonthName(month) + " " + year); // Create a new Spreadsheet
    var sheet = ss.getActiveSheet(); // Get the first sheet
    
    clearSheet(sheet);

    setWeekdaysHeader(sheet);

    const daysInMonth = new Date(year, month, 0).getDate();
    let firstDayOfWeek = new Date(year, month - 1, 1).getDay();
    let date = 1;
    let lastRow;

    for (let row = 3; date <= daysInMonth; row += 3) {
        for (let col = 8; col >= 2; col--) {
            const cell = sheet.getRange(row, col);
            const curColDayOfWeek = 7 - (col - 1);

            if (firstDayOfWeek === curColDayOfWeek && date === 1) {
                cell.setValue(date++).setHorizontalAlignment('center');
                firstDayOfWeek = -1;
            } else if (date > 1 && date <= daysInMonth) {
                cell.setValue(date++).setHorizontalAlignment('center');
            } else {
                cell.setBackground("#CCCCCC");
            }

            cell.setFontSize(12);
            setMonthColor(month, cell);
            
            // Setting up rows and columns properties
            sheet.getRange(row + 1, col).setBackground("#FFFFFF").setFontSize(8);
            sheet.getRange(row + 2, col).setBackground("#e8dddc").setFontSize(8);
            sheet.setRowHeight(row, 20);
            sheet.setRowHeight(row + 1, 80);
            sheet.setRowHeight(row + 2, 20);
        }

        // Set up the sum formula for this row
        const sumFormulaForThisRow = `=SUM(B${row + 2}:H${row + 2})`;
        sheet.getRange(row + 2, 9).setFormula(sumFormulaForThisRow);
        lastRow = row + 2;
    }

    // Post-loop adjustments
    sheet.getRange(2, 2, lastRow - 1, 7).setBorder(true, true, true, true, false, true);
    sheet.deleteColumns(10, sheet.getMaxColumns() - 9);
    sheet.deleteRows(lastRow + 1, sheet.getMaxRows() - lastRow);

    const sumCell = sheet.getRange(lastRow + 1, 8);
    const formula = `=SUM(I5:I${lastRow})`;
    sumCell.setFormula(formula).setFontSize(12).setBackground("#E0E0E0");

    // Setting up the remaining backgrounds
    sheet.getRange(2, 9, lastRow).setBackground("#E0E0E0");
    sheet.getRange(lastRow + 1, 2, 1, 7).setBackground("#E0E0E0");

    // Remove the first column and set the month name
    sheet.deleteColumn(1);
    const monthName = getMonthName(month);
    sheet.getRange("A1").setValue(monthName).setFontWeight("bold").setFontSize(14).setHorizontalAlignment('center');
}

// Function to set up the trigger
function setupTrigger() {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1;  // JavaScript months are 0-based
  ScriptApp.newTrigger('autoCreateMonthlySheet')
    .timeBased()
    .onMonthDay(1)  // Fire on the 1st of every month
    .atHour(1)  // At 1:00 AM
    .create();
}

// Function to automatically create the monthly sheet and send an email
function autoCreateMonthlySheet() {
  var today = new Date();
  var nextMonth = today.getMonth() + 1;  // Modified this line
  var year = (nextMonth === 12) ? today.getFullYear() + 1 : today.getFullYear();
  var month = (nextMonth === 12) ? 1 : nextMonth + 1;  // Modified this line

  createMonthlySheet(year, month);

  // Send an email notification
  var emailAddress = "your_mail@mail.com";
  var subject = "New Monthly Sheet Created";
  var message = `The new monthly sheet for ${getMonthName(month)} ${year} has been created and is ready for use.`;
  
  MailApp.sendEmail(emailAddress, subject, message);
}

// Run this function once to set up the trigger
setupTrigger();

