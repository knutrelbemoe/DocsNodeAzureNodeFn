var request = require("request");
var adal = require("adal-node");
var fs = require("fs");

module.exports = async function (context, req) {
    var authorityHostUrl = 'https://login.microsoftonline.com';
    var tenant = ''; //'docsnode.com';
    var fileName = '';
    var folderName = '';
    var userGuidId = '';

    if (req.body && req.body.tenant) {
        tenant = req.body.tenant;
        fileName = req.body.FileName;
        folderName = req.body.FolderName;
        sourceFileName = req.body.sourceFileName;
        userGuidId = req.body.userGuidId;
    }

    var fileExtension = sourceFileName.split('.');
    var fileExt = fileExtension[fileExtension.length - 1];
    fileName = fileName + "." + fileExt;
    var authorityUrl = authorityHostUrl + '/' + tenant;
    var certificate = fs.readFileSync('devcert.pem', {
        encoding: 'utf8'
    });
    var clientId = process.env['Dev-AD-APP-ClientID'];
    var thumbprint = process.env['Dev-Cert-Thumbprint'];

    var authContext = new adal.AuthenticationContext(authorityUrl);

    return new Promise((resolve, reject) => {
        authContext.acquireTokenWithClientCertificate("https://graph.microsoft.com", clientId, certificate, thumbprint, function (err, tokenResponse) {
            if (err) {
                context.log('well that didn\'t work: ' + err.stack);
                context.done();
                return;
            }
            context.log(tokenResponse);
            var accesstoken = tokenResponse.accessToken;
            var options = {
                method: "GET",
                uri: "https://graph.microsoft.com/v1.0/users/" + userGuidId + "/drive/root:/" + folderName + "/" + fileName + "?select=name,id,webUrl",
                headers: {
                    'Authorization': 'Bearer ' + accesstoken,
                    'Accept': 'application/json;odata.metadata=full'
                }
            };
            context.log(options);
            request(options, function (error, res, body) {
                context.log(error);
                context.log(body);
                //var formatJSON = JSON.parse(body);
                //var res = formatJSON.webUrl;
                resolve(body);
                context.res = {
                    body: body || ''
                };
                context.done();
            });
        });
    });
};