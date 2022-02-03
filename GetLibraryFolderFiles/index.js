var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


module.exports = function (context, req) {

    var authorityHostUrl = 'https://login.microsoftonline.com';
    var tenant = ''; //'docsnode.com';
    var resource = '';
    var folderPath = '';
    var accountName = '';
    var tenantName = '';
    if (req.body && req.body.tenant && req.body.SPOUrl) {
        resource = req.body.SPOUrl;
        tenant = req.body.tenant;
        folderPath = req.body.FolderPath;
        accountName = req.body.AccountName;
        tenantName = req.body.TenantName;
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
        var folderProp = JSON.stringify({
            tenantName: tenantName,
            siteUrl: resource,
            folderPath: folderPath,
            emailID: accountName
        });
        var optionsPermission = {
            method: "POST",
            uri: "https://docsnodecore-function.azurewebsites.net/api/CheckFolderPermission?code=Lrj4CxoYzSCV4v5hdHvNZOJa9ck417U6pg5abYFw4sgXdoZts6aAqA==",
            headers: {
                'Content-Type': 'application/json'
            },
            body: folderProp
        };

        request(optionsPermission, function (error, res, data) {
            context.log(error);
            var formatJSON = JSON.parse(data);

            if (formatJSON.hasPermission) {


                var options = {
                    method: "GET",
                    uri: resource + "/sites/DocsNodeAdmin/_api/web/GetFolderByServerRelativeUrl('/sites/DocsNodeAdmin/DocsNodeTemplatesLibrary/" + folderPath + "')?$expand=Files,Folders",
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
            }
            else {
                var folderNotAccess = { folderAccess: "Access Denined" };
                context.res = {
                    body: folderNotAccess
                };
                context.done();
            }
        });
    });
};