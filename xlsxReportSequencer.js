const ExcelJS = require('exceljs');
const { attendanceMatrix} = require('./attendanceMatrix.js');

async function reportSequencer(auth, searchTerms, template, outputFile, startDate, endDate) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Report');

    let currentRow = 1;

    for (const element of template.elements) {
        switch (element.type) {
            case 'attendanceMatrix':
                const matrixWorkbook = await attendanceMatrix(auth, searchTerms, startDate, endDate);
                const sourceWorksheet = matrixWorkbook.getWorksheet(1);
                const targetStartingCell = { row: currentRow, col: element.startCol || 1 }; // Changed variable name

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

    const filename = `report_${new Date().toISOString().replace(/[:.]/g, '-')}.xlsx`;
    await workbook.xlsx.writeFile(filename);
    console.log(`Report saved to "${filename}"`);
}

function replicateContentToSheetPosition(sourceWorksheet, targetWorksheet, targetStartingCell) {
    let lastRow = 0;

    // Set rightToLeft property on the target worksheet
    targetWorksheet.views = [{ rightToLeft: true }];

    sourceWorksheet.eachRow({ includeEmpty: true }, (sourceRow, sourceRowIndex) => {
        sourceRow.eachCell({ includeEmpty: true }, (sourceCell, sourceColIndex) => {
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
};
