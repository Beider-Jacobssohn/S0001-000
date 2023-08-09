//Current entrypoint (runProgram)

const fs = require('fs');
const authorize = require('./googleAuthorization.js');
const {attendanceMatrix} = require('./attendanceMatrix.js');
const {generateReportFromTemplate} = require('./xlsxReportSequencer.js');
const templates = JSON.parse(fs.readFileSync('./Data/templates.json', 'utf8'));
const institutionsData = JSON.parse(fs.readFileSync('./Data/institutions.json', 'utf8'));
// const studentsData = JSON.parse(fs.readFileSync('./Data/students.json', 'utf8'));

async function runProgram() {
    //logging:
    // console.log(institutionsData);
    const auth = await authorize();

    for (const institutionObj of institutionsData) {
        // console.log("------+ obj ", institutionObj)
        for (const institutionKey in institutionObj) {
            // console.log("-----= key ", institutionKey)
            const institution = institutionObj[institutionKey];
            // console.log(institution);  // This should now print the institution keys (e.g., "123456789")
            const teachers = institution.teachers;
            // console.log("------ ",teachers);

            for (const teacherID in teachers) {
                const teacher = teachers[teacherID];
                const courses = teacher.courses;

                for (const courseName in courses) {
                    console.log("courseName: ",courseName)
                    const course = courses[courseName];
                    console.log(course)
                    console.log(course["studentIDs"])

                    if (!course || !course.studentIDs || course.studentIDs.length === 0) {
                        console.warn(`Skipping course "${courseName}" due to missing student data.`);
                        continue;
                    }

                    // Replace student IDs with their corresponding search terms
                    // const students = course.students.map(studentID => studentsData[studentID]?.searchTerms);
                    // const students = course.studentIDs.flatMap(studentID => {
                    //     console.log(studentsData);
                    //     console.log("-----* ", studentsData[studentID]["searchTerms"])
                    //     return studentsData[studentID]["searchTerms"];
                    // });
                    const studentIDs = course.studentIDs;
                    if (!studentIDs || studentIDs.length === 0) {
                        console.warn(`Skipping course "${courseName}" due to missing search terms.`);
                        continue;
                    }

                    // Generate the attendance matrix
                    try {
                        (`Generating attendance matrix for course "${courseName}" of teacher "${teacherID}" in institution "${institutionKey}"...`);

                        const attendanceMatrixWorkbook = await attendanceMatrix(
                            auth,
                            studentIDs,
                            new Date('2023-01-01T00:00:00Z'),
                            new Date()
                        );

                        // Save the attendance matrix workbook to a file
                        const outputFileName = `./Data/Output/attendanceMatrix-${institutionKey}-${teacherID}-${courseName}-${new Date().toISOString().replace(/:/g, '-')}.xlsx`;
                        await attendanceMatrixWorkbook.xlsx.writeFile(outputFileName);
                        console.log(`-----0 Attendance matrix saved to "${outputFileName}"`);
                    } catch (error) {
                        console.error(`-----( Failed to generate attendance matrix for course "${courseName}" of teacher "${teacherID}" in institution "${institutionKey}":`, error);
                        continue;
                    }

                    // Generate the report
                    try {
                        console.log(`-----9 Generating report for course "${courseName}" of teacher "${teacherID}" in institution "${institutionKey}"...`);
                        await generateReportFromTemplate(
                            auth,
                            studentIDs,  // Changed from "students" to "studentIDs"
                            templates.conservatoryReport, // use the conservatoryReport template object
                            `./Data/Output/report-${institutionKey}-${teacherID}-${courseName}-${new Date().toISOString().replace(/:/g, '-')}.xlsx`,
                            new Date('2023-01-01T00:00:00Z'),
                            new Date()
                        );
                    } catch (error) {
                        console.error(`Failed to generate report for course "${courseName}" of teacher "${teacherID}" in institution "${institutionKey}":`, error);
                    }
                }
            }
        }
    }
}

runProgram().then(r => console.log(r));