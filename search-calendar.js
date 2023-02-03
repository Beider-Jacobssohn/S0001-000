const { google } = require('googleapis');

const searchCalendar = async (auth, searchTerms) => {
    const calendar = google.calendar({ version: 'v3', auth });
    const now = new Date();
    const last30Days = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

    const events = await calendar.events.list({
        calendarId: 'primary',
        timeMin: last30Days.toISOString(),
        timeMax: now.toISOString(),
        singleEvents: true,
        orderBy: 'startTime',
        maxResults: 1000,
        fields: 'items(id,summary,start,end)',
    });

    // Search for events that match the search terms
    const matchingEvents = events.data.items.filter(event => {
        return searchTerms.some(term => {
            return (
                event.summary.toLowerCase().includes(term.toLowerCase()) ||
                (event.description && event.description.toLowerCase().includes(term.toLowerCase()))
            );
        });
    });

    // Return the matching events
    return matchingEvents;
}

module.exports = searchCalendar;
