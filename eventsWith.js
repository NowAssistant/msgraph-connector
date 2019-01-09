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

        if (
            people.statusCode === 200 && people.body.value && people.body.value.length > 0 &&
            events.statusCode === 200 && events.body.value && events.body.value.length > 0
        ) {
            const matches = [];

            for (let i = 0; i < people.body.value.length; i++) {
                for (let j = 0; j < events.body.value.length; j++) {
                    if (
                        people.body.value[i].userPrincipalName ===
                        events.body.value[j].organizer.emailAddress.address
                    ) {
                        matches.push(events.body.value[j]);
                        continue;
                    }

                    for (let k = 0; k < events.body.value[j].attendees.length; k++) {
                        if (
                            people.body.value[i].userPrincipalName ===
                            events.body.value[j].attendees[k].emailAddress.address
                        ) {
                            matches.push(events.body.value[j]);
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
