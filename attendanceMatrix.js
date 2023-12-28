const ExcelJS = require('exceljs');
const fs = require('fs');
const searchCalendar = require("./search-calendar");
const AMstudentsData = JSON.parse(fs.readFileSync('./Data/students.json', 'utf8'));

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

async function createAttendanceMatrix(events, studentIDs, startingCell = { row: 1, col: 1 }) {
    const uniqueDates = getUniqueDates(events);
    const workbook = new ExcelJS.Workbook();

    // Write unique dates on the X-axis for each month's worksheet
    events.forEach(event => {
        const startDate = new Date(event.start.dateTime || event.start.date);
        const month = startDate.getMonth();
        const year = startDate.getFullYear();
        const formattedDate = formatDate(startDate);

        const worksheet = getWorksheetForMonth(workbook, month, year);
        writeUniqueDatesToWorksheet(uniqueDates, month, year, worksheet, startingCell);

        const summary = event.summary.toLowerCase();  // Let's make it lowercase once for efficiency
        // console.log(`Event Summary: ${event.summary}`);  // Log the event summary

        studentIDs.forEach((studentID, index) => {
            const student = AMstudentsData[studentID];
            const row = startingCell.row + index + 1;
            const col = startingCell.col;

            // Writing student names on the Y-axis
            if (student) {
                worksheet.getCell(row, col).value = `${student.name.he} (${studentID}, ${student.phoneNumbers.join(", ")})`;
            } else {
                worksheet.getCell(row, col).value = `Unknown (${studentID})`;
            }

            // for (i in student)
            // {
            //     console.log("----+-",i)
            // }
            // console.log(`STUDENT NAME: ${student.name}\nSTUDENT[1]: ${student[1]} \n sUMMARY: ${summary}`)

            // Checking if the student has a matching event and marking it
            if (student && summary.includes(student.name.en.toLowerCase())) {
                // console.log(`Match found for student ${student.name} in event: ${event.summary}`);  // Log when a match is found
                let dateColumn = startingCell.col + 1;
                while (worksheet.getCell(startingCell.row, dateColumn).value !== formattedDate) {
                    dateColumn++;
                }

                const cell = worksheet.getCell(row, dateColumn);
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

async function attendanceMatrix(auth, studentIDs, startDate, endDate, startingCell = { row: 1, col: 1 }) {
    // Convert studentIDs to searchTerms for fetching events
    const searchTerms = studentIDs.flatMap(studentID => AMstudentsData[studentID]?.searchTerms || []);

    const matchingEvents = await searchCalendar(auth, searchTerms, startDate, endDate);

    return await createAttendanceMatrix(matchingEvents, studentIDs, startingCell);
}

async function attendanceMatrixForTemplate(auth, searchTerms, startDate, endDate) {
    // Fetch matching events using the searchCalendar function
    const matchingEvents = await searchCalendar(auth, searchTerms, startDate, endDate);

    // Create the attendance matrix using the createAttendanceMatrix function
    const matrixWorkbook = await createAttendanceMatrix(matchingEvents, searchTerms);

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


