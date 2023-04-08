const { google } = require('googleapis');

const searchCalendar = async (auth, searchTerms, startDate, endDate) => {
    const calendar = google.calendar({ version: 'v3', auth });

    const events = await calendar.events.list({
        calendarId: 'primary',
        timeMin: startDate.toISOString(),
        timeMax: endDate.toISOString(),
        singleEvents: true,
        orderBy: 'startTime',
        maxResults: 1000,
        fields: 'items(id,summary,start,end)',
    });



    // Search for events that match the search terms
    // Return matching events
    return events.data.items.filter(event => {
        return searchTerms.some(term => {
            return (
                event.summary.toLowerCase().includes(term.toLowerCase()) ||
                (event.description && event.description.toLowerCase().includes(term.toLowerCase()))
            );
        });
    })
}
module.exports = searchCalendar;
