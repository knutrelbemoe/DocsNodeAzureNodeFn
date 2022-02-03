// CreateListItem
var request = require("request");
var adal = require("adal-node");
var fs = require("fs");
// const util = require('util');
var ctx = require("./contextInfo")

module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var reqDigest = '';
	var documentLibrary = '';
	var documentLibraryUrl = '';
	var siteUrl = '';
	var title = "New";
	var pinnedType = '';
	if (req.body && req.body.tenant && req.body.SPOUrl) {
		resource = req.body.SPOUrl;
		tenant = req.body.tenant;
		reqDigest = req.body.ReqDigest;
		documentLibrary = req.body.DocumentLibrary;
		documentLibraryUrl = req.body.DocumentLibraryUrl;
		siteUrl = req.body.SiteUrl;
		pinnedType = req.body.PinnedType;
	}

	var authorityUrl = authorityHostUrl + '/' + tenant;

	//var resource = 'https://docsnode.sharepoint.com';


	var certificate = fs.readFileSync('devcert.pem', {
		encoding: 'utf8'
	});
	var clientId = process.env['Dev-AD-APP-ClientID'];
	var thumbprint = process.env['Dev-Cert-Thumbprint'];
	var accessToken = '';
	// const asyncFunction = util.promisify(ctx.getReqDigest);

	/////////////// section for pin item check /////////////////////

	//	var chkPinnedUrl= resource+"/_api/web/lists/getbytitle('DocsNodePinnedLocations')/items";



	ctx.getReqDigest(context, tenant, resource).then(result => {
		//console.log(accessToken);
		var itemProperties = JSON.stringify({
			__metadata: {
				type: "SP.Data.DocsNodePinnedLocationsListItem"
			},
			Title: title,
			PinnedType: pinnedType,
			DocumentLibrary: documentLibrary,
			DocumentLibraryURL: {
				'__metadata': {
					'type': 'SP.FieldUrlValue'
				},
				'Description': 'Library Url',
				'Url': documentLibraryUrl
			},
			SiteURL: {
				'__metadata': {
					'type': 'SP.FieldUrlValue'
				},
				'Description': 'Site Url',
				'Url': siteUrl
			}
		});


		var options = {
			method: "POST",
			async: false,
			uri: resource + "/_api/web/lists/getbytitle('DocsNodePinnedLocations')/items",
			body: itemProperties,
			headers: {
				'Authorization': 'Bearer ' + result.accessToken,
				'Accept': 'application/json; odata=verbose',
				'Content-Type': 'application/json; odata=verbose',
				//	'X-RequestDigest': result.digestVal,
				'X-HTTP-Method': 'POST'
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



	//	var authContext = new adal.AuthenticationContext(authorityUrl);


	// authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
	// 	if (err) {
	// 		context.log('well that didn\'t work: ' + err.stack);
	// 		context.done();
	// 		return;
	// 	}
	// 	context.log(tokenResponse);

	// //	var accesstoken = tokenResponse.accessToken;
	// 	getReqDigest(tenant,SPOUrl);
	// /*	var options = {
	// 		method: "POST",
	// 		async: false,
	// 		uri: resource + "/_api/contextinfo",
	// 		headers: {
	// 			'Authorization': 'Bearer ' + accesstoken,
	// 			'Accept': 'application/json; odata=verbose',
	// 			'Content-Type': 'application/json; odata=verbose'
	// 		}
	// 	};

	// 	context.log(options);
	// 	request(options, function (error, res, body) {
	// 		context.log(error);
	// 		reqDigest = JSON.parse(body).d.GetContextWebInformation.FormDigestValue;
	//     }); */

	// });

	// var authContextCreate = new adal.AuthenticationContext(authorityUrl);

	// authContextCreate.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
	// 	if (err) {
	// 		context.log('well that didn\'t work: ' + err.stack);
	// 		context.done();
	// 		return;
	// 	}

	// 	context.log(tokenResponse);
	// 	var accesstoken = tokenResponse.accessToken;



	//	});
};
