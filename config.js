const currentUTCDateTime = require('./lib/get-utcdatetime');

module.exports = {
    CACHE: process.env.CACHE === 'true',
    graph: {
        auth: {
            url: process.env.GRAPH_AUTH_ENDPOINT || 'https://login.microsoftonline.com/vtfk.onmicrosoft.com/oauth2/v2.0/token',
            clientId: process.env.GRAPH_AUTH_CLIENT_ID || '123456-1234-1234-123456',
            secret: process.env.GRAPH_AUTH_SECRET || 'wnksdnsjblnsfjb',
            scope: process.env.GRAPH_AUTH_SCOPE || 'https://graph.microsoft.com/.default',
            grantType: process.env.GRAPH_AUTH_GRANT_TYPE || 'client_credentials'
        },
        user: {
            meUrl: process.env.GRAPH_ME_ENDPOINT || 'https://graph.microsoft.com/v1.0/me',
            userUrl: process.env.GRAPH_USERS_ENDPOINT || 'https://graph.microsoft.com/v1.0/users',
            rootUrl: process.env.GRAPH_ORG_ROOT || 'https://graph.microsoft.com/v1.0/sites/root/',
            properties: 'id,userPrincipalName,mail,displayName'
        },
        org: {
            url: process.env.GRAPH_ORG_ENDPOINT || 'https://graph.microsoft.com/v1.0/organization',
            properties: 'id'
        },
        calender: {
            graphUrl: process.env.OUTLOOK_ENDPOINT || 'https://graph.microsoft.com/beta/me/events',
            link: process.env.OUTLOOK_LINK_URL || 'https://outlook.office365.com/owa/?itemid=',
            maxTasks: parseInt(process.env.OUTLOOK_MAX_EVENTS) || 100,
            filter: process.env.OUTLOOK_FILTER || 'start/dateTime ge \'' + currentUTCDateTime() + '\' and start/dateTime le \'' + currentUTCDateTime(process.env.FILTER_CALENDER_DATE_ADD) + '\'',
            orderBy: process.env.OUTLOOK_ORDERBY || 'start/datetime'
        },
    }
};