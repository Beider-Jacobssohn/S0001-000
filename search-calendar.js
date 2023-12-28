const { google } = require('googleapis');

const searchCalendar = async (auth, searchTerms, startDate, endDate) => {
    const calendar = google.calendar({ version: 'v3', auth });

    // flatSearchTerms = searchTerms.flat

    // console.log(typeof(searchTerms[0][0]));
    // console.log(searchTerms[0])

    // console.log("Seach terms: ", searchTerms)

    // Search for events that match the search terms
    const events = await calendar.events.list({
        calendarId: 'primary',
        timeMin: startDate.toISOString(),
        timeMax: endDate.toISOString(),
        singleEvents: true,
        orderBy: 'startTime',
        maxResults: 1000,
        fields: 'items(id,summary,start,end)',
    });

    // Return matching events
    console.log("@@@", searchTerms)
    return events.data.items.filter(event => {
        return searchTerms.some(term => {
            // console.log("Term: ",term[0])
            return (
                event.summary.toLowerCase().includes(term.toLowerCase()) ||
                (event.description && event.description.toLowerCase().includes(term.toLowerCase()))
            );
        });
    })
}
module.exports = searchCalendar;
