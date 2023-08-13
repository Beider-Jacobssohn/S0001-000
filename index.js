const fs = require('fs');
const authorize = require('./googleAuthorization.js');
const { generateReportFromTemplate } = require('./xlsxReportSequencer.js');
const templates = JSON.parse(fs.readFileSync('./Data/templates.json', 'utf8'));
const institutionsData = JSON.parse(fs.readFileSync('./Data/institutions.json', 'utf8'));
const ExcelJS = require('exceljs');

async function runProgram() {
    const auth = await authorize();
    const workbook = new ExcelJS.Workbook(); // Create the workbook outside the loops

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

                for (const courseName in courses) {
                    console.log("courseName: ", courseName);
                    const course = courses[courseName];

                    if (!course || !course.studentIDs || course.studentIDs.length === 0) {
                        console.warn(`Skipping course "${courseName}" due to missing student data.`);
                        continue;
                    }

                    const studentIDs = course.studentIDs;
                    if (!studentIDs || studentIDs.length === 0) {
                        console.warn(`Skipping course "${courseName}" due to missing search terms.`);
                        continue;
                    }

                    try {
                        console.log(`Generating report for course "${courseName}" of teacher "${teacherID}" in institution "${institutionKey}"...`);
                        await generateReportFromTemplate(
                            auth,
                            studentIDs,
                            institutionKey,
                            teacherID,
                            courseName,
                            templates.conservatoryReport,
                            workbook, // Pass the workbook
                            courseName, // Use the course name as the sheet name
                            startDate,
                            endDate
                        );

                        // Extract year and month from startDate for the filename
                        const year = startDate.getFullYear();
                        const month = String(startDate.getMonth() + 1).padStart(2, '0');

                        // Get the teacher's English name
                        const teacherEnglishName = institution?.[institutionKey]?.teachers?.[teacherID]?.name?.en;

                        // Formulate the base filename
                        let baseFilename = `./Data/Output/${teacherEnglishName}-${year}-${month}`;
                        let finalFilename = `${baseFilename}.xlsx`;
                        let counter = 1;

                        // Check if the file exists, and if so, append a sequential number
                        while (fs.existsSync(finalFilename)) {
                            finalFilename = `${baseFilename}-${counter}.xlsx`;
                            counter++;
                        }

                        await workbook.xlsx.writeFile(finalFilename);
                        console.log(`++++++Report saved to "${finalFilename}"`);

                    } catch (error) {
                        console.error(`Failed to generate report for course "${courseName}" of teacher "${teacherID}" in institution "${institutionKey}":`, error);
                    }
                }
            }
        }
    }
}

runProgram().then(r => console.log(r));

