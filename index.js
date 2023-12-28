const fs = require('fs');
const authorize = require('./googleAuthorization.js');
const { generateReportFromTemplate } = require('./xlsxReportSequencer.js');
const templates = JSON.parse(fs.readFileSync('./Data/templates.json', 'utf8'));
const institutionsData = JSON.parse(fs.readFileSync('./Data/institutions.json', 'utf8'));
const ExcelJS = require('exceljs');

async function runProgram() {
    const auth = await authorize();

    // Define the startDate and endDate constants
    const startDate = new Date('2023-01-01T00:00:00Z');
    const endDate = new Date();

    for (const institutionObj of institutionsData) {
        for (const institutionKey in institutionObj) {
            const institution = institutionObj[institutionKey];
            const teachers = institution.teachers;

            for (const teacherID in teachers) {
                const teacher = teachers[teacherID];
                const courses = teacher.courses;

                let currentDate = new Date(startDate.getTime());

                while (currentDate <= endDate) {
                    // Extract year and month from currentDate for the worksheet name
                    const year = currentDate.getFullYear();
                    const month = String(currentDate.getMonth() + 1).padStart(2, '0');

                    // Create the workbook at the start of the month loop
                    const workbook = new ExcelJS.Workbook();

                    for (const courseName in courses) {
                        console.log("courseName: ", courseName);
                        const course = courses[courseName];

                        if (!course || !course.studentIDs || course.studentIDs.length === 0) {
                            console.warn(`Skipping course "${courseName}" due to missing student data.`);
                            continue;
                        }

                        const studentIDs = course.studentIDs;
                        if (!studentIDs || studentIDs.length === 0) {
                            console.warn(`Skipping course "${courseName}" due to lack of students.`);
                            continue;
                        }

                        try {
                            console.log(`Generating report for course "${courseName}" of teacher "${teacherID}" in institution "${institutionKey}"...`);

                            // Append the year and month to the course name
                            const sheetName = `${courseName}-${year}-${month}`;

                            await generateReportFromTemplate(
                                auth,
                                studentIDs,
                                institutionKey,
                                teacherID,
                                courseName,
                                templates.conservatoryReport,
                                workbook, // Pass the workbook
                                sheetName, // Use the course name with date as the sheet name
                                currentDate,
                                endDate
                            );

                        } catch (error) {
                            console.error(`Failed to generate report for course "${courseName}" of teacher "${teacherID}" in institution "${institutionKey}":`, error);
                        }
                    }

                    // Get the teacher's first and last names, or use a default value
                    const teacherFirstName = institutionObj[institutionKey]?.teachers?.[teacherID]?.firstName || 'defaultFirstName';
                    const teacherLastName = institutionObj[institutionKey]?.teachers?.[teacherID]?.lastName || 'defaultLastName';

                    // Concatenate the first and last names with a space in between
                    const teacherFullName = `${teacherFirstName} ${teacherLastName}`;

                    // Formulate the base filename
                    let baseFilename = `./Data/Output/${teacherFullName}-${year}-${month}`;
                    let finalFilename = `${baseFilename}.xlsx`;
                    let counter = 1;

                    // Check if the file exists, and if so, append a sequential number
                    while (fs.existsSync(finalFilename)) {
                        finalFilename = `${baseFilename}-${counter}.xlsx`;
                        counter++;
                    }

                    await workbook.xlsx.writeFile(finalFilename);
                    console.log(`++++++Report saved to "${finalFilename}"`);

                    // Update currentDate to the next month
                    currentDate.setMonth(currentDate.getMonth() + 1);
                }
            }
        }
    }
}

runProgram().then(r => console.log(r));