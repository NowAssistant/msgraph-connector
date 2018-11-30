'use strict';

const got = require('got');
const moment = require('moment');
const api = require('./api');

const fields = [
    'subject',
    'body',
    'bodyPreview',
    'organizer',
    'attendees',
    'start',
    'end',
    'location',
    'isCancelled',
    'webLink',
    'onlineMeetingUrl',
    'createdDateTime',
    'lastModifiedDateTime',
    'reminderMinutesBeforeStart',
    'isReminderOn',
    'responseRequested',
    'responseStatus'
];

const dateAscending = (a, b) => {
    a = new Date(a.date);
    b = new Date(b.date);

    return a < b ? -1 : (a > b ? 1 : 0);
};

module.exports = async function (activity) {
    try {
        api.initialize(activity);

        const response = await api('/me/events?$select=' + fields.join(','));

        if (response.statusCode === 200 && response.body.value && response.body.value.length > 0) {
            const today = new Date();

            const items = [];

            for (let i = 0; i < response.body.value.length; i++) {
                const item = await convertItem(response.body.value[i]);
                const eventDate = new Date(item.date);

                if (today.setHours(0, 0, 0, 0) === eventDate.setHours(0, 0, 0, 0)) {
                    items.push(item);
                }
            }

            if (items.length > 0) {
                activity.Response.Data.items = items.sort(dateAscending);
            } else {
                activity.Response.Data.items = [];
            }
        } else {
            activity.Response.Data = {
                statusCode: response.statusCode,
                message: 'Bad request or no events returned',
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

    return activity; // cloud connector support

    async function convertItem(_item) {
        const item = _item;

        item.date = new Date(_item.start.dateTime).toISOString();

        const _duration = moment.duration(
            moment(_item.end.dateTime)
                .diff(
                    moment(_item.start.dateTime)
                )
        );

        let duration = '';

        if (_duration._data.years > 0) {
            duration += _duration._data.years + 'y ';
        }

        if (_duration._data.months > 0) {
            duration += _duration._data.months + 'mo ';
        }

        if (_duration._data.days > 0) {
            duration += _duration._data.days + 'd ';
        }

        if (_duration._data.hours > 0) {
            duration += _duration._data.hours + 'h ';
        }

        if (_duration._data.minutes > 0) {
            duration += _duration._data.minutes + 'm ';
        }

        item.duration = duration;

        if (item.location && item.location.coordinates) {
            item.location.link =
                'https://www.google.com/maps/search/?api=1&query=' +
                item.location.coordinates.latitude + ',' +
                item.location.coordinates.longitude;
        }

        item.organizer.photo = await fetchPhoto(_item.organizer.emailAddress.address);

        for (let i = 0; i < _item.attendees.length; i++) {
            item.attendees[i].photo = await fetchPhoto(_item.attendees[i].emailAddress.address);
        }

        item.showDetails = false;

        return item;
    }

    async function fetchPhoto(account) {
        const endpoint =
            activity.Context.connector.endpoint + '/users/' +
            account + '/photos/48x48/$value';

        try {
            const response = await got(endpoint, {
                headers: {
                    Authorization: 'Bearer ' + activity.Context.connector.token
                },
                encoding: null
            });

            if (response.statusCode === 200) {
                return 'data:' + response.headers['content-type'] + ';base64,' +
                    new Buffer(response.body).toString('base64');
            }
        } catch (err) {
            return null;
        }
    }
};
