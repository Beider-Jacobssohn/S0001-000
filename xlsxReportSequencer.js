const ExcelJS = require('exceljs');
const { attendanceMatrix} = require('./attendanceMatrix.js');
const fs = require("fs");
const institutionsData = JSON.parse(fs.readFileSync('./Data/institutions.json', 'utf8'));

function insertText(worksheet, options = {}, currentRow) {
    const col = options.startCol || 1;
    const cell = worksheet.getCell(currentRow, col);
    cell.value = options.value || '';
}

function insertTeacherName(sheet, options, currentRow) {
    const cell = sheet.getCell(currentRow, options.startCol || 1);
    cell.value = options.teacherName; // "teacherName" is just an example, replace it with the actual property name from options
}

function insertInstitutionName(sheet, options, currentRow) {
    const cell = sheet.getCell(currentRow, options.startCol || 1);
    cell.value = options.institutionName; // "institutionName" is just an example, replace it with the actual property name from options
}


async function reportSequencer(auth, searchTerms, institutionKey, teacherID, courseName, template, outputFile, startDate, endDate) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Report');

    let currentRow = 1;

    // Extract data from institutions.json using institutionKey, teacherID, and courseName
    const institution = institutionsData.find(inst => inst[institutionKey]);
    const teacher = institution?.[institutionKey]?.teachers?.[teacherID];

    if (institution?.[institutionKey]?.institutionName) {
        insertInstitutionName(sheet, { institutionName: institution[institutionKey].institutionName }, currentRow);
        currentRow++;
    }
    if (teacher?.firstName && teacher?.lastName) {
        // Concatenate firstName and lastName for RTL languages
        const fullName = `${teacher.lastName} ${teacher.firstName}`;
        insertTeacherName(sheet, { teacherName: fullName }, currentRow);
        currentRow++;
    }
    if (courseName) {
        insertText(sheet, { value: `Course: ${courseName}` }, currentRow);
        currentRow++;
    }

    for (const element of template.elements) {
        switch (element.type) {
            case 'attendanceMatrix':
                const matrixWorkbook = await attendanceMatrix(auth, searchTerms, startDate, endDate);
                const sourceWorksheet = matrixWorkbook.getWorksheet(1);
                const targetStartingCell = {row: currentRow, col: element.startCol || 1};

                currentRow = replicateContentToSheetPosition(sourceWorksheet, sheet, targetStartingCell);
                break;
            case 'teacherName':
                insertTeacherName(sheet, element.options, currentRow);
                currentRow++;
                break;
            case 'institutionName':
                insertInstitutionName(sheet, element.options, currentRow);
                currentRow++;
                break;
            case 'text':
                insertText(sheet, element.options, currentRow);
                currentRow++;
                break;
            default:
                console.warn(`Unknown element type: ${element.type}`);
        }
    }

    const filename = `./Data/Output/report_${new Date().toISOString().replace(/[:.]/g, '-')}.xlsx`;
    try {
        await workbook.xlsx.writeFile(filename);
        console.log(`++++++Report saved to "${filename}"`);
    } catch (error) {
        console.error(`Failed to save the report to "${filename}":`, error);
    }
}

function replicateContentToSheetPosition(sourceWorksheet, targetWorksheet, targetStartingCell) {
    let lastRow = 0;

    if (!sourceWorksheet || !sourceWorksheet.eachRow) {
        console.warn('No valid data found for replication in the report.');
        return lastRow;
    }

    // Set rightToLeft property on the target worksheet
    targetWorksheet.views = [{rightToLeft: true}];

    sourceWorksheet.eachRow({includeEmpty: true}, (sourceRow, sourceRowIndex) => {
        sourceRow.eachCell({includeEmpty: true}, (sourceCell, sourceColIndex) => {
            const targetRowIndex = targetStartingCell.row + sourceRowIndex - 1;
            const targetColIndex = targetStartingCell.col + sourceColIndex - 1;

            const targetCell = targetWorksheet.getCell(targetRowIndex, targetColIndex);
            targetCell.value = sourceCell.value;
            targetCell.style = sourceCell.style;

            lastRow = Math.max(lastRow, targetRowIndex);
        });
    });

    return lastRow + 1;
}

    module.exports = {
        generateReportFromTemplate: reportSequencer
    }
