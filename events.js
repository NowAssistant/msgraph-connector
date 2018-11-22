const api = require('./api');
const got = require('got');
const moment = require('moment');

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
 
        configure_range();
 
        activity.Response.Data.items = [];
 
        for (let i = 0; i < response.body.value.length; i++) {
            const item = await convert_item(response.body.value[i]);

            if (!skip(i, response.body.value.length, new Date(item.date))) {
                activity.Response.Data.items.push(item);
            }
        }

        activity.Response.Data.items.sort(dateAscending);
    } catch (error) {
        var m = error.message;
 
        if (error.stack) {
            m = m + ': ' + error.stack;
        }
 
        activity.Response.ErrorCode = (error.response && error.response.statusCode) || 500;
        activity.Response.Data = {
            ErrorText: m
        };
    }

    return activity; // cloud connector support
 
    function configure_range() {
        if (activity.Request.Query.startDate) {
            startDate = convert_date(activity.Request.Query.startDate);
        }
 
        if (activity.Request.Query.endDate) {
            endDate = convert_date(activity.Request.Query.endDate);
        }
 
        if (activity.Request.Query.page && activity.Request.Query.pageSize) {
            action = 'firstpage';
            page = parseInt(activity.Request.Query.page);
            pageSize = parseInt(activity.Request.Query.pageSize);
 
            if (activity.Request.Data &&
                activity.Request.Data.args &&
                activity.Request.Data.args.atAgentAction == 'nextpage') {
                action = 'nextpage';
                page = parseInt(activity.Request.Data.args._page) || 2;
                pageSize = parseInt(activity.Request.Data.args._pageSize) || 20;
            }
        } else if (activity.Request.Query.pageSize) {
            pageSize = parseInt(activity.Request.Query.pageSize);
        } else {
            pageSize = 10;
        }
    }
 
    async function convert_item(_item) {
        const item = _item;
 
        item.date = new Date(_item.start.dateTime).toISOString();

        /*item.organizer.photo = await fetch_photo(
            _item.organizer.emailAddress.address
        );*/

        const _duration = moment.duration(
            moment(_item.end.dateTime)
                .diff(
                    moment(_item.start.dateTime)
                )
        );

        let duration = '';

        if (_duration._data.years > 0) duration += _duration._data.years + 'y ';
        if (_duration._data.months > 0) duration += _duration._data.months + 'mo ';
        if (_duration._data.days > 0) duration += _duration._data.days + 'd ';
        if (_duration._data.hours > 0) duration += _duration._data.hours + 'h ';
        if (_duration._data.minutes > 0) duration += _duration._data.minutes + 'm ';

        item.duration = duration;

        const start = new Date(_item.start.dateTime);

        item.time = start.getHours() + ':' + 
            (start.getMinutes() < 10 ? '0' + start.getMinutes() : start.getMinutes());

        if (item.location && item.location.coordinates) {
            item.location.link = 
                'https://www.google.com/maps/search/?api=1&query=' + 
                item.location.coordinates.latitude + ',' + 
                item.location.coordinates.longitude;
        }

        item.attendeesCount = item.attendees.length;

        if (item.responseRequested && item.responseStatus.response == 'accepted') {
            item.responseStatus.string = 'Accepted ' + moment(item.responseStatus.time).fromNow();
        }

        item.date_readable = moment(item.date).format('MM/DD/YY');

        /*for (let i = 0; i < _item.attendees.length; i++) {
            item.attendees[i].photo = await fetch_photo(
                item.attendees[i].emailAddress.address
            );
        }*/

        return item;
    }

    async function fetch_photo(account) {
        const endpoint = 
            activity.Context.connector.endpoint + '/users/' + 
            account + '/photos/48x48/$value';
 
        try {
            const response = await got(endpoint, {
                headers: {
                    'Authorization': 'Bearer ' + activity.Context.connector.token
                },
            });

            if (response.statusCode == 200) {
                console.log(response);
                const data = new Buffer.from(response.body, 'binary');

                return 'data:image/jpeg;base64,' + data.toString('base64');
            }
        } catch (err) {
            console.log(err);
            return null;
        }
    }
};

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
 
function convert_date(date) {
    return new Date(
        date.substring(0, 4),
        date.substring(4, 6) - 1,
        date.substring(6, 8)
    );
}
 
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