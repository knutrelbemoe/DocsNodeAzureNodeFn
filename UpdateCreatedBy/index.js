// UpdateCreatedBy
var request = require("request");
var adal = require("adal-node");
var fs = require("fs");
var csomapi = require('csom-node');


module.exports = function (context, req) {

	var authorityHostUrl = 'https://login.microsoftonline.com';
	var tenant = ''; //'docsnode.com';
	var resource = '';
	var siteUrl = '';
	var listName = '';
	var itemID = '';
	var userID = '';

	if (req.body && req.body.tenant && req.body.SPOUrl && req.body.ItemID) {
		resource = req.body.SPOUrl;
		tenant = req.body.tenant;
		siteUrl = req.body.SiteUrl;
		listName = req.body.ListName;
		itemID = req.body.ItemID;
		userID = req.body.UserID;
	}

	var authorityUrl = authorityHostUrl + '/' + tenant;

	//var resource = 'https://docsnode.sharepoint.com';


	var certificate = fs.readFileSync('devcert.pem', {
		encoding: 'utf8'
	});
	var clientId = process.env['Dev-AD-APP-ClientID'];
	var thumbprint = process.env['Dev-Cert-Thumbprint'];

	var authContext = new adal.AuthenticationContext(authorityUrl);

	var authCtx = new AuthenticationContext(resource);

	authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
		if (err) {
			context.log('well that didn\'t work: ' + err.stack);
			context.done();
			return;
		}
		context.log(tokenResponse);

		authCtx.appAccessToken = tokenResponse.accessToken;
		var accesstoken = tokenResponse.accessToken;

		var ctx = new SP.ClientContext("/");  //set root web
		authCtx.setAuthenticationCookie(ctx);  //authenticate

		//retrieve SP.Web client object
		//var web = ctx.get_web();
		//	var list = ctx.get_web().get_lists().getByTitle(listName);
		//	var item = list.getItemById(itemID);
		//	item.set_item("Editor", userID);
		//	item.update();

		//ctx.load(web);

		var web = ctx.get_web();
		ctx.load(web);
		ctx.executeQueryAsync(function () {
			azureContext.log(web.get_title());
			azureContext.res = { body: "Success!" };
			azureContext.done();
		},
			function (sender, args) {
				azureContext.log('An error occured: ' + args.get_message());
				azureContext.res = { status: 500, body: "Error!" };
				azureContext.done();
			});

	});
};