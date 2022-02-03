var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


module.exports = function (context, req) {

    var authorityHostUrl = 'https://login.microsoftonline.com';
    var tenant = ''; //'docsnode.com';
    var resource = '';
    var team = '';
    if (req.body && req.body.tenant && req.body.SPOUrl) {
        resource = req.body.SPOUrl;
        tenant = req.body.tenant;
        teamUrl = req.body.Team;
    }

    var authorityUrl = authorityHostUrl + '/' + tenant;
    var teamAllLibraryUrl = teamUrl + "/_api/Web/Lists?$select=EntityTypeName,title&$filter=(BaseTemplate eq 101)&$expand=Folders";
    context.log(teamUrl);

    //var resource = 'https://docsnode.sharepoint.com';


    var certificate = fs.readFileSync('devcert.pem', {
        encoding: 'utf8'
    });
    var clientId = process.env['Dev-AD-APP-ClientID'];
    var thumbprint = process.env['Dev-Cert-Thumbprint'];

    var authContext = new adal.AuthenticationContext(authorityUrl);

    authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
        if (err) {
            context.log('well that didn\'t work: ' + err.stack);
            context.done();
            return;
        }
        context.log(tokenResponse);

        var accesstoken = tokenResponse.accessToken;

        var options = {
            method: "GET",
            uri: teamAllLibraryUrl,
            headers: {
                'Authorization': 'Bearer ' + accesstoken,
                'Accept': 'application/json; odata=verbose',
                'Content-Type': 'application/json; odata=verbose'
            }
        };


        context.log(options);
        request(options, function (error, res, body) {
            context.log(error);
            context.log(body);
            context.res = {
                body: body || ''
            };
            context.done();
        });
    });
};
