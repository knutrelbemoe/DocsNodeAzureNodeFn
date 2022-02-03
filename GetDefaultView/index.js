var request = require("request");
var adal = require("adal-node");
var fs = require("fs");

module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var userEmail = '';

	if (req.body && req.body.tenant && req.body.SPOUrl && req.body.UserEmail) {
		resource = req.body.SPOUrl;
		tenant = req.body.tenant;
		userEmail = req.body.UserEmail;
	}
	var siteUrl = resource + "/sites/DocsNodeAdmin";
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

		var options = {
			method: "GET",
			uri: siteUrl + "/_api/lists/getbytitle('DocsNodeDefaultView')/items?$filter=UserEmail eq'" + userEmail + "'",
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