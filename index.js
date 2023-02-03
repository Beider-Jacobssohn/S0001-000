const authorize = require('./googleAuthorization.js');
const searchCalendar = require('./search-calendar.js');

async function runProgram() {
    const auth = await authorize();
    const searchTerms = ['meeting', 'appointment'];
    const matchingEvents = await searchCalendar(auth, searchTerms);
    console.log(matchingEvents);
}

runProgram();
