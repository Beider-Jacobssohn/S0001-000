const express = require('express');
const fs = require('fs');
const authorize = require('./googleAuthorization.js');
const { generateReportFromTemplate } = require('./xlsxReportSequencer.js');
const ExcelJS = require('exceljs');

const app = express();
const port = 3002;

app.use(express.json());

const templates = JSON.parse(fs.readFileSync('./Data/templates.json', 'utf8'));
const institutionsData = JSON.parse(fs.readFileSync('./Data/institutions.json', 'utf8'));

app.post('/generate-gcal-report', async (req, res) => {
  try {
    const { startDate, endDate, institution, teacherID } = req.body;
    const auth = await authorize();

    const institutionObj = institutionsData.find(inst => Object.keys(inst)[0] === institution);
    if (!institutionObj) {
      throw new Error('Institution not found');
    }

    const teacher = institutionObj[institution].teachers[teacherID];
    if (!teacher) {
      throw new Error('Teacher not found');
    }

    const courses = teacher.courses;
    let currentDate = new Date(startDate);
    const endDateObj = new Date(endDate);

    while (currentDate <= endDateObj) {
      const year = currentDate.getFullYear();
      const month = String(currentDate.getMonth() + 1).padStart(2, '0');

      const workbook = new ExcelJS.Workbook();

      for (const courseName in courses) {
        console.log("courseName: ", courseName);
        const course = courses[courseName];

        if (!course || !course.studentIDs || course.studentIDs.length === 0) {
          console.warn(`Skipping course "${courseName}" due to missing student data.`);
          continue;
        }

        const studentIDs = course.studentIDs;

        try {
          console.log(`Generating report for course "${courseName}" of teacher "${teacherID}" in institution "${institution}"...`);

          const sheetName = `${courseName}-${year}-${month}`;

          await generateReportFromTemplate(
            auth,
            studentIDs,
            institution,
            teacherID,
            courseName,
            templates.conservatoryReport,
            workbook,
            sheetName,
            currentDate,
            endDateObj
          );

        } catch (error) {
          console.error(`Failed to generate report for course "${courseName}" of teacher "${teacherID}" in institution "${institution}":`, error);
        }
      }

      const teacherFirstName = teacher.firstName || 'defaultFirstName';
      const teacherLastName = teacher.lastName || 'defaultLastName';
      const teacherFullName = `${teacherFirstName} ${teacherLastName}`;

      let baseFilename = `./Data/Output/${teacherFullName}-${year}-${month}`;
      let finalFilename = `${baseFilename}.xlsx`;
      let counter = 1;

      while (fs.existsSync(finalFilename)) {
        finalFilename = `${baseFilename}-${counter}.xlsx`;
        counter++;
      }

      await workbook.xlsx.writeFile(finalFilename);
      console.log(`++++++Report saved to "${finalFilename}"`);

      currentDate.setMonth(currentDate.getMonth() + 1);
    }

    res.json({ message: 'Reports generated successfully' });
  } catch (error) {
    console.error('Error generating report:', error);
    res.status(500).json({ error: 'Failed to generate report', details: error.message });
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
