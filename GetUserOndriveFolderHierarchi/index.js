var request = require("request");
var adal = require("adal-node");
var fs = require("fs");
var ctx = require("./driveId")

module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var UserGuidId = '';
	var ItemID = '';
	if (req.body && req.body.tenant && req.body.ItemID) {
		resource = "https://graph.microsoft.com";
		tenant = req.body.tenant;
		UserGuidId = req.body.UsrGUID;
		ItemID = req.body.ItemID;
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
			//uri: "https://graph.microsoft.com/v1.0/users/"+UserGuidId+"/drives/"+result.driveId+"/root/children",
			url: "https://graph.microsoft.com/v1.0/users/" + UserGuidId + "/drives/" + result.driveId + "/items/" + ItemID + "/children",
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