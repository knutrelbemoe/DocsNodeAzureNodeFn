var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


exports.getDriveId = function getDriveId(context, tenant, resource, UserGuidId) {
	var authorityHostUrl = 'https://login.microsoftonline.com';
	var authorityUrl = authorityHostUrl + '/' + tenant;
	var certificate = fs.readFileSync('devcert.pem', {
		encoding: 'utf8'
	});
	var clientId = process.env['Dev-AD-APP-ClientID'];
	var thumbprint = process.env['Dev-Cert-Thumbprint'];
	var authContext = new adal.AuthenticationContext(authorityUrl);
	var accessToken = '';
	return new Promise((resolve, reject) => {
		return authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint,
			function (err, tokenResponse) {
				if (err) {
					context.log('well that didn\'t work: ' + err.stack);
					context.done();
					return;
				}

				var options = {
					method: "GET",
					uri: "https://graph.microsoft.com/v1.0/users/" + UserGuidId + "/drives",
					headers: {
						'Authorization': 'Bearer ' + tokenResponse.accessToken,
						'Accept': 'application/json;odata.metadata=full'
					}
				};

				context.log(options);
				request(options, function (error, res, body) {
					var formatJSON = JSON.parse(body);
					var res = { accessToken: tokenResponse.accessToken, driveId: formatJSON.value[0].id };
					resolve(res);
				});
			});
	});
}