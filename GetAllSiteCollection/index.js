var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


module.exports = function (context, req) {

    var authorityHostUrl = 'https://login.microsoftonline.com';
    var tenant = ''; //'docsnode.com';
    var resource = '';
    if (req.body && req.body.tenant && req.body.SPOUrl) {
        resource = req.body.SPOUrl;
        tenant = req.body.tenant;
    }

    var authorityUrl = authorityHostUrl + '/' + tenant;

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
        var tenantName = "";
        function splitStr(tenant) {
            // Function to split string 
            var string = tenant.split(".");
            tenantName = string[0];
        }
        splitStr(tenant);
        var options = {
            method: "GET",
            uri: resource + "/_api/search/query?querytext=%27NOT%20Path:https://" + tenantName + "-my.sharepoint.com/personal/*%20contentclass:sts_site%27&selectproperties=%27Title,Path%27&rowLimit=499&TrimDuplicates=false",
            //  uri:"https://graph.microsoft.com/v1.0//users/arijit@docsnode.com/memberOf",
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

//note : check if we can get list of all site collection only where login user has access using single api call.