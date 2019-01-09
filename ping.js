'use strict';

const api = require('./api');

module.exports = async function (activity) {
    try {
        api.initialize(activity);

        const response = await api('/me');

        activity.Response.Data = {
            success: (response && response.statusCode === 200)
        };
    } catch (error) {
        const response = {
            success: false
        };

        let m = error.message;

        if (error.stack) {
            m = m + ': ' + error.stack;
        }

        response.ErrorText = m;
        activity.Response.Data = response;
    }
};
