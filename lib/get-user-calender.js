const {
    CACHE
} = require('../config');

const getOutlookTasks = require('../lib/calender/get-outlook-tasks');
const HTTPError = require('../lib/http-error');

const NodeCache = require('node-cache');
const cache = CACHE ? new NodeCache({
    stdTTL: 3600,
    checkperiod: 120
}) : false;

module.exports = async (context, graphUser, token) => {
    const {
        userPrincipalName,
        onPremisesSamAccountName: samAccountName
    } = graphUser;

    let outlookTasks = []
    try {
        if (token) outlookTasks = await getOutlookTasks(context, token, userPrincipalName);
    } catch (error) {
        context.log.error(['events', 'get-user-calender', userPrincipalName, 'get-outlook-events', 'err', error]);
    }

    try {


        context.log(['tasks', 'get-user-tasks', userPrincipalName, 'tasks', outlookTasks.count]);

        if (cache) {
            cache.set(userPrincipalName, outlookTasks)
        }

        return outlookTasks;
    } catch (err) {
        context.log.error(['tasks', userPrincipalName, 'err', err]);
        throw new HTTPError(err.statusCode || 500, err.message);
    }
}