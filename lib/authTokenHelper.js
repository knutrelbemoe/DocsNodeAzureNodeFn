var adal = require("adal-node");
var fs = require("fs");


module.exports = async (resource, tenant) => {
    var certificate = fs.readFileSync('devcert.pem', {
        encoding: 'utf8'
    });
    var clientId = process.env['Dev-AD-APP-ClientID'];
    var thumbprint = process.env['Dev-Cert-Thumbprint'];
    var authorityHostUrl = 'https://login.microsoftonline.com';
    var authorityUrl = authorityHostUrl + '/' + tenant;

    var authContext = new adal.AuthenticationContext(authorityUrl);


    return new Promise((resolve, reject) => {
        authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, token) {
            if (err) {
                reject(err);
            } else {
                resolve(token);
            }
        });
    });
}