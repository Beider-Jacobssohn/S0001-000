const ExcelJS = require('exceljs');
const searchCalendar = require("./search-calendar");

function getUniqueDates(events) {
    const dates = events.map(event => {
        const startDate = new Date(event.start.dateTime || event.start.date);
        return startDate.toISOString().split('T')[0];
    });
    return [...new Set(dates)].sort();
}

function formatDate(date) {
    const d = new Date(date);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
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

async function createAttendanceMatrix(events, searchTerms, startingCell = { row: 1, col: 1 }) {
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

    return workbook; // Return the workbook containing the attendance matrix
}

async function attendanceMatrix(auth, searchTerms, startDate , endDate, startingCell = { row: 1, col: 1 }) {
    const matchingEvents = await searchCalendar(auth, searchTerms, startDate, endDate);
    return await createAttendanceMatrix(matchingEvents, searchTerms, startingCell);
}

async function attendanceMatrixForTemplate(auth, searchTerms, startDate, endDate) {
    // Fetch matching events using the searchCalendar function
    const matchingEvents = await searchCalendar(auth, searchTerms, startDate, endDate);

    // Create the attendance matrix using the createAttendanceMatrix function
    const matrixWorkbook = await createAttendanceMatrix(matchingEvents, searchTerms);
    const outputFile = `matrixWorkbook-${new Date().toISOString().replace(/:/g, '-')}.xlsx`;
    await matrixWorkbook.xlsx.writeFile(outputFile);
    console.log(`Matrix workbook saved to "${outputFile}"`);


    // Extract the attendance matrix as a 2D array from the matrixWorkbook
    const matrix = [];
    let rowIndex = 0;
    matrixWorkbook.eachSheet((worksheet, sheetId) => {
        worksheet.eachRow((row, rowNumber) => {
            if (sheetId > 1 && rowNumber === 1) return; // Ignore the first row (headers) for all sheets after the first one
            matrix[rowIndex] = row.values.slice(1); // Ignore the first element (it's always empty)
            rowIndex++;
        });
    });

    return matrix;
}

module.exports = {
    attendanceMatrix ,createAttendanceMatrix, attendanceMatrixForTemplate
};


