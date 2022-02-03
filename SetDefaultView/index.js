var request = require("request");
var adal = require("adal-node");
var fs = require("fs");
module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var userEmail = '';
	var folderPath = '';
	var status = '';
	if (req.body && req.body.tenant && req.body.SPOUrl) {
		resource = req.body.SPOUrl;
		tenant = req.body.tenant;
		userEmail = req.body.UserEmail;
		folderPath = req.body.FolderPath;
		status = req.body.Status;
	}

	var siteUrl = resource + "/sites/DocsNodeAdmin";
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
			var formatJSON = JSON.parse(body);
			if (formatJSON.d.results.length > 0) {
				//update existing item
				var itemID = formatJSON.d.results[0].ID;
				authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
					if (err) {
						context.log('well that didn\'t work: ' + err.stack);
						context.done();
						return;
					}
					context.log(tokenResponse);

					var accesstoken = tokenResponse.accessToken;

					var updateProperties = JSON.stringify({
						__metadata: {
							type: "SP.Data.DocsNodeDefaultViewListItem"
						},
						FolderPath: folderPath,
						Active: status
					});

					var optionsDefaultViewUpdate = {
						method: "POST",
						async: false,
						uri: siteUrl + "/_api/web/lists/getbytitle('DocsNodeDefaultView')/items(" + itemID + ")",
						body: updateProperties,
						headers: {
							'Authorization': 'Bearer ' + accesstoken,
							'Accept': 'application/json; odata=verbose',
							'Content-Type': 'application/json; odata=verbose',
							'X-HTTP-Method': 'MERGE',
							'IF-MATCH': '*'
						}
					};
					context.log(optionsDefaultViewUpdate);
					request(optionsDefaultViewUpdate, function (error, res, body) {
						context.log(error);
						context.log(body);
						context.res = {
							body: body || ''
						};
						context.done();
					});
				});
			}

			else {

				//insert new item

				authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
					if (err) {
						context.log('well that didn\'t work: ' + err.stack);
						context.done();
						return;
					}
					context.log(tokenResponse);

					var accesstoken = tokenResponse.accessToken;

					var itemProperties = JSON.stringify({
						__metadata: {
							type: "SP.Data.DocsNodeDefaultViewListItem"
						},
						Title: "New",
						UserEmail: userEmail,
						FolderPath: folderPath,
						Active: true
					});

					var optionsDefaultViewInsert = {
						method: "POST",
						async: false,
						uri: siteUrl + "/_api/web/lists/getbytitle('DocsNodeDefaultView')/items",
						body: itemProperties,
						headers: {
							'Authorization': 'Bearer ' + accesstoken,
							'Accept': 'application/json; odata=verbose',
							'Content-Type': 'application/json; odata=verbose',
							'X-HTTP-Method': 'POST'
						}
					};
					context.log(optionsDefaultViewInsert);
					request(optionsDefaultViewInsert, function (error, res, body) {
						context.log(error);
						context.log(body);
						context.res = {
							body: body || ''
						};
						context.done();
					});
				});
			}
		});
	});
};