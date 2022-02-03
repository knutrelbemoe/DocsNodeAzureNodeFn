var request = require("request");
var adal = require("adal-node");
var fs = require("fs");


module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var sourceSite = '';
	var destSite = '';
	var sourceFileRelUrl = '';
	var sourceFileName = '';
	var destFolderRelUrl = '';
	var fileName = '';

	if (req.body && req.body.tenant && req.body.SPOUrl) {
		resource = req.body.SPOUrl;
		tenant = req.body.tenant;
		destSite = req.body.DestSite;
		destFolderRelUrl = req.body.DestFolderRelUrl;
		fileName = req.body.FileName;
		sourceFileName = req.body.sourceFileName;
	}

	sourceSite = resource + "/sites/docsnodeadmin";
	sourceFileRelUrl = "/sites/docsnodeadmin/DocsNodeTemplatesLibrary/" + sourceFileName;
	var fileExtension = sourceFileName.split('.');
	var fileExt = fileExtension[fileExtension.length - 1];
	fileName = fileName + "." + fileExt;
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
			uri: sourceSite + "/_api/web/GetFileByServerRelativeUrl('" + sourceFileRelUrl + "')/openbinarystream",
			encoding: null,
			headers: {
				'Authorization': 'Bearer ' + accesstoken
			}
		};
		context.log(options);
		request(options, function (error, res, body) {
			context.log(error);
			var sampleBytes = new Uint8Array(body);
			if (sampleBytes.length > 0) {
				// Construct the endpoint.
				var fileCollectionEndpoint = destSite + "/_api/web/getfolderbyserverrelativeurl('" + destFolderRelUrl + "')/files" + "/add(overwrite=false, url='" + fileName + "')";
				options = {
					method: "POST",
					uri: fileCollectionEndpoint,
					body: sampleBytes,
					processData: false,
					headers: {
						'Authorization': 'Bearer ' + accesstoken,
						'Accept': 'application/json; odata=verbose',
						'Content-Type': 'application/json; odata=verbose'
					}
				};
				context.log(options);
				request(options, function (error, res, body) {
					context.log(error);
					var formatJSON = JSON.parse(body);
					if (formatJSON.error === undefined) {

						//	var fileURL = destSite +"_api/web/lists/getbytitle('"+ destFolderRelUrl +"')/Items?$filter=FileLeafRef eq '"+ fileName +"'&$select*";
						options = {
							method: "GET",
							uri: formatJSON.d.ListItemAllFields.__deferred.uri,
							//	encoding: null,
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
					} else {
						context.log(error);
						context.log(body);
						context.res = {
							body: body || ''
						};
						context.done();
					}
				});
			} else {
				context.done();
			}
		});
	});
};