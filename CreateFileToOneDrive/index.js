var request = require("request");
var adal = require("adal-node");
var fs = require("fs");
var ctx = require("./driveId");
var async = require('async');

module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var sourceSite = '';
	var sourceFileRelUrl = '';
	var sourceFileName = '';
	var fileName = '';
	var folderName = '';
	var userGuidId = '';

	if (req.body && req.body.tenant && req.body.SPOUrl) {
		resource = req.body.SPOUrl;
		tenant = req.body.tenant;
		fileName = req.body.FileName;
		sourceFileName = req.body.sourceFileName;
		folderName = req.body.FolderName;
		userGuidId = req.body.userGuidId;
	}

	sourceSite = resource + "/sites/docsnodeadmin";
	sourceFileRelUrl = "/sites/docsnodeadmin/DocsNodeTemplatesLibrary/" + sourceFileName;
	var fileExtension = sourceFileName.split('.');
	var fileExt = fileExtension[fileExtension.length - 1];
	fileName = fileName + "." + fileExt;
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
			uri: sourceSite + "/_api/web/GetFileByServerRelativeUrl('" + sourceFileRelUrl + "')/openbinarystream",
			encoding: null,
			headers: {
				'Authorization': 'Bearer ' + accesstoken
			}
		};
		context.log(options);


		authContext.acquireTokenWithClientCertificate("https://graph.microsoft.com", clientId, certificate, thumbprint, function (err, tokenResponseGraph) {
			if (err) {
				context.log('well that didn\'t work: ' + err.stack);
				context.done();
				return;
			}
			context.log(tokenResponseGraph);

			var accesstokenGraph = tokenResponseGraph.accessToken;

			request(options, function (error, res, body) {
				context.log(error);
				var sampleBytes = new Uint8Array(body);

				if (sampleBytes.length > 0) {
					//var oneDriveEndpoint = "https://graph.microsoft.com/v1.0/users/" + userGuidId + "/drive/items/" + itemID + "/createUploadSession";
					var oneDriveEndpoint = "";
					if (folderName !== "/") {
						oneDriveEndpoint = "https://graph.microsoft.com/v1.0/users/" + userGuidId + "/drive/root:/" + folderName + "/" + fileName + ":/createUploadSession";
					} else {
						oneDriveEndpoint = "https://graph.microsoft.com/v1.0/users/" + userGuidId + "/drive/root:/" + fileName + ":/createUploadSession";
					}
					var optionsUploadFilest = {
						method: "POST",
						//	url: 'https://graph.microsoft.com/v1.0/drive/root:/Attachments/' + onedrive_file + ':/createUploadSession',
						url: oneDriveEndpoint,
						headers: {
							'Authorization': 'Bearer ' + accesstokenGraph,
							'Content-Type': "application/json",
							'Accept': 'application/json;odata.metadata=full'
							// 'Content-Type': mime.getType(file)
						},
						//	data: {"item": {"@odata.type": "microsoft.graph.driveItemUploadableProperties","@microsoft.graph.conflictBehavior": "rename","name": "largefile.docx"}}
					};
					request(optionsUploadFilest, function (error, res, body) {
						uploadFile(JSON.parse(body).uploadUrl, sampleBytes);
						context.log(optionsUploadFilest);
						context.res = {
							body: body || ''
						};
						context.done();
					});
				} else {
					context.done();
				}
			});
		});
	});



	function uploadFile(uploadUrl, f) { // Here, it uploads the file by every chunk.
		async.eachSeries(getparams(f.length), function (st, callback) {
			setTimeout(function () {
				var optionsFilePost = {
					method: "PUT",
					url: uploadUrl,
					headers: {
						'Content-Length': st.clen,
						'Content-Range': st.cr,
					},
					body: f.slice(st.bstart, st.bend + 1),
				};
				request(optionsFilePost, function (error, res, body) {
					context.log(error);
				});
				callback();
			}, st.stime);
		});
	}



	function getparams(allsize) {
		//	var allsize = fs.statSync(file).size;
		var sep = allsize < (60 * 1024 * 1024) ? allsize : (60 * 1024 * 1024) - 1;
		var ar = [];
		for (var i = 0; i < allsize; i += sep) {
			var bstart = i;
			var bend = i + sep - 1 < allsize ? i + sep - 1 : allsize - 1;
			var cr = 'bytes ' + bstart + '-' + bend + '/' + allsize;
			var clen = bend != allsize - 1 ? sep : allsize - i;
			var stime = allsize < (60 * 1024 * 1024) ? 5000 : 10000;
			ar.push({
				bstart: bstart,
				bend: bend,
				cr: cr,
				clen: clen,
				stime: stime,
			});
		}
		return ar;
	}
};