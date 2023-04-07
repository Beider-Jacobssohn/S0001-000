const ExcelJS = require('exceljs');

function getUniqueDates(events) {
    const dates = events.map(event => {
        const startDate = new Date(event.start.dateTime || event.start.date);
        return startDate.toISOString().split('T')[0];
    });
    return [...new Set(dates)].sort();
}

function getWorksheetForMonth(workbook, month, year) {
    const monthName = new Date(year, month).toLocaleString('default', { month: 'long' });
    const sheetName = `${monthName} ${year}`;
    let worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
        worksheet = workbook.addWorksheet(sheetName);
        worksheet.views = [{ rightToLeft: true }];
    }

    return worksheet;
}

function writeUniqueDatesToWorksheet(uniqueDates, month, year, worksheet, startingCell) {
    let column = startingCell.col + 1;
    uniqueDates.forEach((date) => {
        const dateObj = new Date(date);
        if (dateObj.getMonth() === month && dateObj.getFullYear() === year) {
            worksheet.getCell(startingCell.row, column).value = date;
            column++;
        }
    });
}


async function writeEventsToXlsx(events, searchTerms, outputFile, startingCell = { row: 1, col: 1 }) {
    const uniqueDates = getUniqueDates(events);
    const workbook = new ExcelJS.Workbook();

    // Write unique dates on the X-axis for each month's worksheet
    events.forEach(event => {
        const startDate = new Date(event.start.dateTime || event.start.date);
        const month = startDate.getMonth();
        const year = startDate.getFullYear();

        const worksheet = getWorksheetForMonth(workbook, month, year);
        writeUniqueDatesToWorksheet(uniqueDates, month, year, worksheet, startingCell);

        // Write search terms on the Y-axis (starting from row 2)
        searchTerms.forEach((term, index) => {
            const row = startingCell.row + index + 1;
            const col = startingCell.col;

            worksheet.getCell(row, col).value = term;
        });

        // Write matching events
        const summary = event.summary;
        const formattedDate = formatDate(startDate);

        searchTerms.forEach((term, termIndex) => {
            if (summary.toLowerCase().includes(term.toLowerCase())) {
                let dateColumn = startingCell.col + 1;
                while (worksheet.getCell(startingCell.row, dateColumn).value !== formattedDate) {
                    dateColumn++;
                }

                const cell = worksheet.getCell(termIndex + startingCell.row + 1, dateColumn);
                cell.value = '✓';

                cell.alignment = {
                    horizontal: 'center',
                    vertical: 'middle',
                };
            }
        });
    });

    await workbook.xlsx.writeFile(outputFile);
}

module.exports = writeEventsToXlsx;
