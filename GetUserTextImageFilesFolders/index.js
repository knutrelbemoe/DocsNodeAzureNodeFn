var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


module.exports = function (context, req) {

    var authorityHostUrl = 'https://login.microsoftonline.com';
    var tenant = ''; //'docsnode.com';
    var resource = '';
    var folderPath='';
    var title = '';
    if (req.body && req.body.tenant && req.body.SPOUrl) {
        resource = req.body.SPOUrl;
        tenant = req.body.tenant;
        folderPath=req.body.FolderPath;
        title=req.body.Title;
    }

    var authorityUrl = authorityHostUrl + '/' + tenant;

    //var resource = 'https://docsnode.sharepoint.com';


    var certificate = fs.readFileSync(__dirname + '/devcert.pem', {
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

        var optionsImagesiteURL = {
            method: "GET",
            uri: resource+ "/sites/DocsNodeAdmin/_api/web/lists/getbytitle('DocsNodeConfiguration')/items?$filter=ConfigAssestTitle eq '"+title+"'",
            headers: {
                'Authorization': 'Bearer ' + accesstoken,
                'Accept': 'application/json; odata=verbose',
                'Content-Type': 'application/json; odata=verbose'
            },
        };

        request(optionsImagesiteURL, function (error, res, body) {
            context.log(error);
            var formatJSON = JSON.parse(body);
            var configlistGuid = formatJSON.d.results[0].ConfigSourceListGUID;
            var configSiteUrl = formatJSON.d.results[0].ConfigSourceListPath;
            var configListName = formatJSON.d.results[0].ConfigSourceList;
            var configSourceList = formatJSON.d.results[0].ConfigSourceDisplayListName;
            var configListRelUrl = configSiteUrl + "/" + configSourceList + "/" + folderPath;

            if(configListName === "Sample Org Assets"){
            var options = {
                method: "GET",
                uri: resource + "/" +configSiteUrl + "/_api/web/GetFolderByServerRelativeUrl('"+configListRelUrl+"')?$expand=Files,Folders",
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
        else if(configListName === "DocsNodeText"){

            // logic for folder structure for SP list

            var options = {
                method: "GET",
                uri: resource + "/" +configSiteUrl + "/_api/web/GetFolderByServerRelativeUrl('"+configListRelUrl+"')?$expand=Files,Folders",
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
        });
 
    });
};