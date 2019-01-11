'use strict';

const moment = require('moment');
const api = require('./api');

const urlRegex = /[-a-zA-Z0-9@:%_\+.~#?&//=]{2,256}\.[a-z]{2,4}\b(\/[-a-zA-Z0-9@:%_\+.~#?&//=]*)?/gi;

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
                const item = convertItem(response.body.value[i]);
                const eventDate = new Date(item.date);

                if (today.setHours(0, 0, 0, 0) === eventDate.setHours(0, 0, 0, 0)) {
                    items.push(item);
                }
            }

            if (items.length > 0) {
                activity.Response.Data.items = items.sort(dateAscending);
            } else {
                activity.Response.Data = {
                    items: [],
                    message: 'No events found for current date'
                };
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
};

function convertItem(_item) {
    const item = _item;

    item.date = new Date(_item.start.dateTime).toISOString();

    const _duration = moment.duration(
        moment(_item.end.dateTime)
            .diff(
                moment(_item.start.dateTime)));

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
        duration += _duration._data.minutes + 'm';
    }

    item.duration = duration.trim();

    if (item.location && item.location.coordinates) {
        item.location.link =
            'https://www.google.com/maps/search/?api=1&query=' +
            item.location.coordinates.latitude + ',' +
            item.location.coordinates.longitude;
    } else if (item.location && !item.onlineMeetingUrl) {
        const url = parseUrl(item.location.displayName);

        if (url !== null) {
            item.onlineMeetingUrl = url;
            item.location = null;
        }
    }

    if (!item.onlineMeetingUrl) {
        const url = parseUrl(item.bodyPreview);

        if (url !== null) {
            item.onlineMeetingUrl = url;
        }
    }

    // Disable avatars until workable solution found

    /* const basePhotoUri = 'https://outlook.office.com/owa/service.svc/s/GetPersonaPhoto?email=';
    const photoSize = '&size=HR64x64';

    item.organizer.photo = basePhotoUri + _item.organizer.emailAddress.address + photoSize;

    for (let i = 0; i < _item.attendees.length; i++) {
        item.attendees[i].photo = basePhotoUri + _item.attendees[i].emailAddress.address + photoSize;
    } */

    item.showDetails = false;

    return item;
}

function parseUrl(text) {
    if (text.search(urlRegex) !== -1) {
        let url = text.substring(text.search(urlRegex), text.length);

        if (url.indexOf(' ') !== -1) {
            url = url.substring(0, url.indexOf(' '));
        }

        if (!url.match(/^[a-zA-Z]+:\/\//)) {
            url = 'https://' + url;
        }

        return url;
    }

    return null;
}
