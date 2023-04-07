const authorize = require('./googleAuthorization.js');
const readline = require('readline');
const { generateReportFromTemplate } = require('./xlsxReportSequencer.js');
const { attendanceMatrix } = require('./attendanceMatrix.js');
const templates = require('./templates.json');
const fs = require('fs');


const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
});

async function promptForSearchTerms() {
    return new Promise((resolve) => {
        let countdown = 0
        const countdownInterval = setInterval(() => {
            console.log(`You have ${countdown} seconds left to enter search terms.`);
            countdown--;
            if (countdown < 0) {
                clearInterval(countdownInterval);
            }
        }, 1000);

        rl.question('Enter search terms separated by commas: ', (input) => {
            clearInterval(countdownInterval);

            if (input.trim() === '') {
                console.log("No search terms provided or time's up. Using default search terms from searchTerms.json.");
                const defaultSearchTerms = JSON.parse(fs.readFileSync('searchTerms.json', 'utf8'));
                resolve(defaultSearchTerms);
            } else {
                const searchTerms = input.split(',').map(term => term.trim());
                resolve(searchTerms);
            }
        });

        setTimeout(() => {
            rl.write(null, { ctrl: true, name: 'u' }); // Clear the input line
            rl.close(); // Close the readline interface
            console.log("Time's up. Using default search terms from searchTerms.json.");
            const defaultSearchTerms = JSON.parse(fs.readFileSync('searchTerms.json', 'utf8'));
            resolve(defaultSearchTerms); // Use default search terms from searchTerms.json
        }, 15000);
    });
}

async function runProgram() {
    const auth = await authorize();
    const searchTerms = await promptForSearchTerms();
    const attendanceMatrixWorkbook = await attendanceMatrix(auth, searchTerms, new Date('2023-01-01T00:00:00Z'), new Date());

    // Save the attendance matrix workbook to a file
    const outputFileName = 'attendanceMatrix' + '-' + new Date().toISOString().replace(/:/g, '-') + '.xlsx';
    await attendanceMatrixWorkbook.xlsx.writeFile(outputFileName);
    console.log(`Attendance matrix saved to "${outputFileName}"`);

    // Generate the report
    await generateReportFromTemplate(auth,
        searchTerms,
        templates.basicMatrix,
        'generateReportFromTemplate.xlsx',
        new Date('2023-01-01T00:00:00Z'),
        new Date());
}

runProgram();

