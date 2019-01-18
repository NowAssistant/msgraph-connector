'use strict';

const api = require('./api');

module.exports = async (activity) => {
    try {
        api.initialize(activity);

        const response = await api('/me/memberOf');

        if (response.statusCode === 200 && response.body.value && response.body.value.length > 0) {
            activity.Response.Data.items = [];

            for (let i = 0; i < response.body.value.length; i++) {
                activity.Response.Data.items.push(convertItem(response.body.value[i]));
            }
        } else {
            activity.Response.Data = {
                statusCode: response.statusCode,
                message: 'Bad request or no group memberships found',
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
    return {
        id: _item.id,
        title: _item.displayName,
        description: _item.description,
        email: _item.mail,
        date: _item.createdDateTime
    };
}
