// var request = require("request");
// var adal = require("adal-node");
// var fs = require("fs");


// module.exports = function (context, req) {

//     var authorityHostUrl = 'https://login.microsoftonline.com';
//     var tenant = ''; //'docsnode.com';
//     var resource = '';
//     var UsrGuid='';
//     if (req.body && req.body.tenant) {
//         resource = "https://graph.microsoft.com";
//         tenant = req.body.tenant;
//         UsrGuid=req.body.UsrGUID;
//     }

//     var authorityUrl = authorityHostUrl + '/' + tenant;

//     //var resource = 'https://docsnode.sharepoint.com';


//     var certificate = fs.readFileSync('devcert.pem', {
//         encoding: 'utf8'
//     });
//     var clientId = process.env['Dev-AD-APP-ClientID'];
//     var thumbprint = process.env['Dev-Cert-Thumbprint'];

//     var authContext = new adal.AuthenticationContext(authorityUrl);

//     authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
//         if (err) {
//             context.log('well that didn\'t work: ' + err.stack);
//             context.done();
//             return;
//         }
//         context.log(tokenResponse);

//         var accesstoken = tokenResponse.accessToken;

//         var options = {
//             method: "GET",
//             uri: "https://graph.microsoft.com/v1.0/users/"+ UsrGuid+"/drives",
//             headers: {
//                 'Authorization': 'Bearer ' + accesstoken,
//                 'Accept': 'application/json;odata.metadata=full'
//             }
//         };


//         context.log(options);
//         request(options, function (error, res, body) {
//             context.log(error);
//             var formatJSON = JSON.parse(body);
//             var driveID=formatJSON.value[0].id;
//             if (formatJSON.error === undefined) 
//             {
//                 options = {
//                     method: "GET",
//                     uri: "https://graph.microsoft.com/v1.0/users/"+UsrGuid+"/drives/"+driveID+"/root/children",
//                     processData: false,
//                     async:false,
//                     headers: {
//                         'Authorization': 'Bearer ' + accesstoken,
//                         'Accept': 'application/json;odata.metadata=full'
//                     }
//                 };
// //context.log(options);
//                 request(options, function (error, res, body) {
//                     context.log(error);
//                     var formatJSON = JSON.parse(body);
//                     context.log(body);
//                     context.res = {
//                         body: body || ''
//                     };
//                     context.done();
//                 });
//             }
//             context.log(body);
//             context.res = {
//                 body: body || ''
//             };
//             context.done();
//         });
//     });
// };













var request = require("request");
var adal = require("adal-node");
var fs = require("fs");
var ctx = require("./driveId")

module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var UserGuidId = '';
	if (req.body && req.body.tenant) {
		resource = "https://graph.microsoft.com";
		tenant = req.body.tenant;
		UserGuidId = req.body.UsrGUID;
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

	ctx.getDriveId(context, tenant, resource, UserGuidId).then(result => {
		//console.log(accessToken);
		var options = {
			method: "GET",
			async: false,
			uri: "https://graph.microsoft.com/v1.0/users/" + UserGuidId + "/drives/" + result.driveId + "/root/children",
			headers: {
				'Authorization': 'Bearer ' + result.accessToken,
				'Accept': 'application/json;odata.metadata=full'
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
