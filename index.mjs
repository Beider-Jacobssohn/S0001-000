import authorize from './googleAuthorization.mjs';
import searchCalendar from './search-calendar.mjs';

async function runProgram() {
    const auth = await authorize();
    const searchTerms = ['meeting', 'appointment', 'Yahli'];
    const matchingEvents = await searchCalendar(auth, searchTerms);
    console.log(matchingEvents);
}

runProgram();