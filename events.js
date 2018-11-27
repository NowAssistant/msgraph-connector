'use strict';

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

let startDate = null;
let endDate = null;

// eslint-disable-next-line no-unused-vars
let action = null;
let page = null;
let pageSize = null;

module.exports = async function (activity) {
    try {
        api.initialize(activity);

        const response = await api('/me/events?$select=' + fields.join(','));

        if (response.statusCode === 200 && response.body.value && response.body.value.length > 0) {
            configureRange();

            activity.Response.Data.items = [];

            for (let i = 0; i < response.body.value.length; i++) {
                const item = await convertItem(response.body.value[i]);

                if (!skip(i, response.body.value.length, new Date(item.date))) {
                    activity.Response.Data.items.push(item);
                }
            }

            activity.Response.Data.items.sort(dateAscending);
        } else {
            activity.Response.Data = {
                statusCode: response.statusCode,
                message: 'Bad request or no events returned'
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

    function convertItem(_item) {
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

        //TODO change to Cisco room link when response property name is known
        if (item.location && item.location.coordinates) {
            item.location.link =
                'https://www.google.com/maps/search/?api=1&query=' +
                item.location.coordinates.latitude + ',' +
                item.location.coordinates.longitude;
        }

        //TODO add Cisco webex link when response property name is known

        const basePhotoUri = 'https://outlook.office.com/owa/service.svc/s/GetPersonaPhoto?email=';
        const photoSize = '&size=HR64x64';

        item.organizer.photo = basePhotoUri + _item.organizer.emailAddress.address + photoSize;

        for (let i = 0; i < _item.attendees.length; i++) {
            item.attendees[i].photo = basePhotoUri + _item.attendees[i].emailAddress.address + photoSize;
        }

        item.showDetails = false;

        return item;
    }

    function configureRange() {
        if (activity.Request.Query.startDate) {
            startDate = convertDate(activity.Request.Query.startDate);
        }

        if (activity.Request.Query.endDate) {
            endDate = convertDate(activity.Request.Query.endDate);
        }

        if (activity.Request.Query.page && activity.Request.Query.pageSize) {
            action = 'firstpage';
            page = parseInt(activity.Request.Query.page, 10);
            pageSize = parseInt(activity.Request.Query.pageSize, 10);

            if (activity.Request.Data &&
                activity.Request.Data.args &&
                activity.Request.Data.args.atAgentAction === 'nextpage') {
                action = 'nextpage';
                page = parseInt(activity.Request.Data.args._page, 10) || 2;
                pageSize = parseInt(activity.Request.Data.args._pageSize, 10) || 20;
            }
        } else if (activity.Request.Query.pageSize) {
            pageSize = parseInt(activity.Request.Query.pageSize, 10);
        } else {
            pageSize = 10;
        }
    }
};

function skip(i, length, date) {
    if (startDate && endDate) {
        return date < startDate || date > endDate;
    } else if (startDate) {
        return date < startDate;
    } else if (endDate) {
        return date > endDate;
    } else if (page && pageSize) {
        const startItem = Math.max(page - 1, 0) * pageSize;

        let endItem = startItem + pageSize;

        if (endItem > length) {
            endItem = length;
        }

        return i < startItem || i >= endItem;
    } else if (pageSize) {
        return i > pageSize - 1;
    } else {
        return false;
    }
}

function convertDate(date) {
    return new Date(
        date.substring(0, 4),
        date.substring(4, 6) - 1,
        date.substring(6, 8)
    );
}
