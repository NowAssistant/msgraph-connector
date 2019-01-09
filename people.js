'use strict';

const api = require('./api');

module.exports = async (activity) => {
    try {
        api.initialize(activity);

        let action = 'firstpage';
        let page = parseInt(activity.Request.Query.page, 10) || 1;
        let pageSize = parseInt(activity.Request.Query.pageSize, 10) || 20;

        if (
            activity.Request.Data &&
            activity.Request.Data.args &&
            activity.Request.Data.args.atAgentAction === 'nextpage'
        ) {
            page = parseInt(activity.Request.Data.args._page, 10) || 2;
            pageSize = parseInt(activity.Request.Data.args._pageSize, 10) || 20;
            action = 'nextpage';
        }

        if (page < 0) {
            page = 1;
        }

        if (pageSize < 1 || pageSize > 99) {
            pageSize = 20;
        }

        const response = await api('/me/people/?$search=' + activity.Request.Query.query);

        if (response.statusCode === 200 && response.body.value && response.body.value.length > 0) {
            activity.Response.Data._action = action;
            activity.Response.Data._page = page;
            activity.Response.Data._pageSize = pageSize;

            activity.Response.Data.items = [];

            const startItem = Math.max(page - 1, 0) * pageSize;
            let endItem = startItem + pageSize;

            if (endItem > response.body.value.length) {
                endItem = response.body.value.length;
            }

            for (let i = startItem; i < endItem; i++) {
                activity.Response.Data.items.push(
                    convertItem(response.body.value[i])
                );
            }
        } else {
            activity.Response.Data = {
                statusCode: response.statusCode,
                message: 'Bad request or no people found',
                items: []
            };
        }
    } catch (error) {
        let m = error.message;

        if (error.stack) {
            m = m + ': ' + error.stack;
        }

        activity.Response.ErrorCode =
            (error.response && error.response.statusCode) || 500;

        activity.Response.Data = {
            ErrorText: m
        };
    }
};

function convertItem(_item) {
    return {
        id: _item.id,
        name: _item.displayName,
        email: _item.userPrincipalName
    };
}
