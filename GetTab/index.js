var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = 'https://graph.microsoft.com';
	var channelID = '';
	var teamID = '';

	if (req.body && req.body.tenant) {
		tenant = req.body.tenant;
		channelID = req.body.ChannelID;
		teamID = req.body.TeamID;
	}

	var authorityUrl = authorityHostUrl + '/' + tenant;


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
			//https://graph.microsoft.com/beta/teams/c0d0ce63-57d7-4321-be52-32e81980357c/channels/19:af3e9385527f411fbb7ae0465d89068e@thread.skype/tabs
			uri: "https://graph.microsoft.com/beta/teams/" + teamID + "/channels/" + channelID + "/tabs?$filter=teamsAppId eq 'com.microsoft.teamspace.tab.files.sharepoint'",
			headers: {
				'Authorization': 'Bearer ' + accesstoken,
				'Accept': 'application/json;odata.metadata=full'
			}
		};


		context.log(options);
		request(options, function (error, res, body) {
			context.log(error);
			context.log(body);
			var formatJSON = JSON.parse(body);
			if (formatJSON["@odata.count"] > 0) {
				context.res = {
					body: body || ''
				};
				context.done();
			}
			else {
				context.res = {
					body: body || ''
				};
				context.done();
			}

		});
	});
};