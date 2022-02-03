var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var siteUrl = '';
	var folderRelPath = ''; //(e.g : Shared Documents/test)
	var channelID = '';
	var teamID = '';
	var folderType = '';
	var TeamSPUrl = "";

	if (req.body && req.body.tenant && req.body.SPOUrl) {
		resource = req.body.SPOUrl;
		tenant = req.body.tenant;
		folderRelPath = req.body.FolderRelPath;
		channelID = req.body.ChannelID;
		teamID = req.body.TeamID;
		folderType = req.body.FolderType;
		TeamSPUrl = req.body.SiteUrl;
	}

	var authorityUrl = authorityHostUrl + '/' + tenant;

	//var resource = 'https://docsnode.sharepoint.com';


	var certificate = fs.readFileSync('devcert.pem', {
		encoding: 'utf8'
	});
	var clientId = process.env['Dev-AD-APP-ClientID'];
	var thumbprint = process.env['Dev-Cert-Thumbprint'];

	var authContext = new adal.AuthenticationContext(authorityUrl);

	if (folderType === "#microsoft.graph.channel") {

		authContext.acquireTokenWithClientCertificate("https://graph.microsoft.com", clientId, certificate, thumbprint, function (err, tokenResponse) {
			if (err) {
				context.log('well that didn\'t work: ' + err.stack);
				context.done();
				return;
			}
			context.log(tokenResponse);

			var accesstoken = tokenResponse.accessToken;

			var optionsChannel = {
				method: "GET",
				//https://graph.microsoft.com/beta/teams/bb021a6a-2089-466e-ad87-8a5439c42696/channels/19:a27dbb83c2ce49a9b409257f2cc258ee@thread.skype/filesFolder
				uri: "https://graph.microsoft.com/beta/teams/" + teamID + "/channels/" + channelID + "/filesFolder",
				headers: {
					'Authorization': 'Bearer ' + accesstoken,
					'Accept': 'application/json;odata.metadata=full'
				}
			};


			context.log(optionsChannel);
			request(optionsChannel, function (error, res, body) {
				context.log(error);
				context.log(body);
				var formatJSON = JSON.parse(body);
				if (!formatJSON.error) {
					var webUrl = formatJSON["webUrl"];

					//Get team Name form Team URl string
					var splitString = webUrl.split("/");
					var teamName = splitString[4];
					var relPath = "/sites/" + teamName + "/" + folderRelPath;
					siteUrl = resource + "/sites/" + teamName;

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
							uri: siteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + relPath + "')/folders",
							async: false,
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
				} else {
					context.log(error);
					context.log(body);
					context.res = {
						body: body || ''
					};
					context.done();
				}
			});
		});
	} else if (folderType === "sharepoint.folder" || folderType === "#microsoft.graph.channel.private") {

		var relPath = folderRelPath;

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
				uri: TeamSPUrl + "/_api/web/GetFolderByServerRelativeUrl('" + relPath + "')/folders",
				async: false,
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
	} else if (folderType === "tab.sharepoint.folder") {


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
				uri: TeamSPUrl + "/_api/web/GetFolderByServerRelativeUrl('" + folderRelPath + "')/folders",
				async: false,
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
	}
};