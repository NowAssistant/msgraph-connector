'use strict';

const api = require('./api');

const dateAscending = (a, b) => {
    a = new Date(a.start.dateTime);
    b = new Date(b.start.dateTime);

    return a < b ? -1 : (a > b ? 1 : 0);
};

module.exports = async (activity) => {
    try {
        api.initialize(activity);

        const people = await api('/me/people?$search=' + activity.Request.Query.query);
        const events = await api('/me/events');

        const isPeople = (people.statusCode === 200) && people.body.value && (people.body.value.length > 0);
        const isEvents = (events.statusCode === 200) && events.body.value && (events.body.value.length > 0);

        if (isPeople && isEvents) {
            const matches = [];

            for (let i = 0; i < people.body.value.length; i++) {
                for (let j = 0; j < events.body.value.length; j++) {
                    const person = people.body.value[i];
                    const event = events.body.value[j];

                    if (person.userPrincipalName === event.organizer.emailAddress.address) {
                        matches.push(event);
                        continue;
                    }

                    for (let k = 0; k < events.body.value[j].attendees.length; k++) {
                        if (person.userPrincipalName === event.attendees[k].emailAddress.address) {
                            matches.push(event);
                            break;
                        }
                    }
                }
            }

            if (matches.length > 0) {
                activity.Response.Data.items = matches.sort(dateAscending);
            } else {
                activity.Response.Data = {
                    items: [],
                    message: 'No events found with the users returned by search'
                };
            }
        } else {
            activity.Response.Data = {
                statusCodes: {
                    people: people.statusCode,
                    events: events.statusCode
                },
                message: 'Bad request or no people/events found',
                items: []
            };
        }
    } catch (error) {
        let m = error.message;

        if (error.stack) {
            m = m + ': ' + error.stack;
        }

        activity.Response.ErrorCode = (error.response && error.response.statusCode) || 500;
        activity.Response.Data = {
            ErrorText: m
        };
    }
};
