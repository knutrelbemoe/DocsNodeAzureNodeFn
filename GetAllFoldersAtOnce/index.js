// GetAllSubSites
var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var siteUrl = '';
	var libName = '';
	var teamName = '';

	if (req.body && req.body.tenant && req.body.SPOUrl && req.body.TeamName) {
		resource = req.body.SPOUrl;
		tenant = req.body.tenant;
		teamName = req.body.TeamName;
		libName = req.body.LibName;
	}
	siteUrl = resource + "/sites/" + teamName;
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
			//	uri: siteUrl + "/_api/Web/GetSubwebsFilteredForCurrentUser(nWebTemplateFilter=-1)?$filter=(WebTemplate ne 'APP')",
			uri: siteUrl + "/_api/Web/lists/getbytitle('" + libName + "')/items?$expand=Folder&$select=ID,Title,FileLeafRef,Folder/ServerRelativeUrl,FileSystemObjectType",
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