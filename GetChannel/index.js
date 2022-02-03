var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var teamName = '';
	var teamID = '';
	var teamUrl = '';
	var libInternalName = '';

	if (req.body && req.body.tenant && req.body.SPOUrl) {
		resource = req.body.SPOUrl;
		tenant = req.body.tenant;
		teamUrl = req.body.TeamURL;
		teamID = req.body.TeamID;
		libInternalName = req.body.LibInternalName;
	}
	//Get team Name form Team URl string
	var splitString = teamUrl.split("/");
	var teamName = splitString[splitString.length - 1]

	var authorityUrl = authorityHostUrl + '/' + tenant;

	//var resource = 'https://docsnode.sharepoint.com';


	var certificate = fs.readFileSync('devcert.pem', {
		encoding: 'utf8'
	});
	var clientId = process.env['Dev-AD-APP-ClientID'];
	var thumbprint = process.env['Dev-Cert-Thumbprint'];

	var authContext = new adal.AuthenticationContext(authorityUrl);

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
			uri: "https://graph.microsoft.com/beta/teams/" + teamID + "/channels",
			headers: {
				'Authorization': 'Bearer ' + accesstoken,
				'Accept': 'application/json;odata.metadata=full'
			}
		};


		context.log(optionsChannel);
		request(optionsChannel, function (error, res, body) {
			context.log(error);
			context.log(body);
			var formatChannel = JSON.parse(body);
			var reqFormatChannel = formatChannel.value;
			var copyFormatChannel = reqFormatChannel;

			//Get All folders of the top level library using rest api

			authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
				if (err) {
					context.log('well that didn\'t work: ' + err.stack);
					context.done();
					return;
				}
				context.log(tokenResponse);

				var accesstoken = tokenResponse.accessToken;

				var optionsFolder = {
					method: "GET",
					///_api/web/GetFolderByServerRelativeUrl('/sites/DocsNodeAdmin/DocsNodeTemplatesLibrary/"+folderPath+"')?$expand=Files,Folders"
					//	uri: siteUrl + "/_api/Web/GetSubwebsFilteredForCurrentUser(nWebTemplateFilter=-1)?$filter=(WebTemplate ne 'APP')",
					uri: teamUrl + "/_api/web/GetFolderByServerRelativeUrl('/sites/" + teamName + "/" + libInternalName + "/')?$expand=Folders&$filter(Folders.displayName neq 'Forms')",
					headers: {
						'Authorization': 'Bearer ' + accesstoken,
						'Accept': 'application/json; odata=verbose',
						'Content-Type': 'application/json; odata=verbose'
					}
				};

				context.log(optionsFolder);
				request(optionsFolder, function (error, res, body) {
					context.log(error);
					context.log(body);
					var formatFolder = JSON.parse(body);
					var reqFormatFolder = formatFolder.d.Folders.results;

					//Compare items in formatChannel and formatFolder JSON to return final JSON Body

					var objJSON = {} // empty Object
					var key = 'Data';
					objJSON[key] = []; // empty Array, which you can push() values into
					// Loop thorough all folder
					for (let index = 0; index < reqFormatFolder.length; index++) {
						const folderName = reqFormatFolder[index].Name;
						const folderRelPath = reqFormatFolder[index].ServerRelativeUrl;
						//Forms Folder excluded 
						if (folderName === "Forms") {
							continue;
						}

						var folderExist = false;
						reqFormatChannel.filter(function (channel) {
							if (channel.displayName === folderName) {
								folderExist = true;
							}
						});
						if (!folderExist) {
							copyFormatChannel.push({
								displayName: folderName,
								webUrl: folderRelPath,
								id: reqFormatFolder[index].UniqueId,
								"@odata.type": "sharepoint.folder"
							});

						}
					}
					context.res = {
						body: copyFormatChannel || ''
					};
					context.done();
				});
			});

			//context.done();
		});
	});
};