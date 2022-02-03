var request = require("request");
var adal = require("adal-node");
var fs = require("fs");
var listItemEntityType = "SP.Data.DocsNodePinnedLocationsListItem";


exports.updateAuthor = function updateAuthor(context, tenant, resource, userID) {
	var authorityHostUrl = 'https://login.microsoftonline.com';
	var authorityUrl = authorityHostUrl + '/' + tenant;

	//var resource = 'https://docsnode.sharepoint.com';


	var certificate = fs.readFileSync('devcert.pem', {
		encoding: 'utf8'
	});
	var clientId = process.env['Dev-AD-APP-ClientID'];
	var thumbprint = process.env['Dev-Cert-Thumbprint'];

	var authContext = new adal.AuthenticationContext(authorityUrl);

	var accessToken = '';

	// return 1;

	return new Promise((resolve, reject) => {
		return authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint,
			function (err, tokenResponse) {
				if (err) {
					context.log('well that didn\'t work: ' + err.stack);
					context.done();
					return;
				}

				var options = {
					method: "POST",
					//   uri: "https://docsnode.sharepoint.com/_api/web/lists/getByTitle('DocsNodePinnedLocations')/Fields(guid'1df5e554-ec7e-46a6-901d-d85a3881cb18')",
					uri: "https://docsnode.sharepoint.com/_api/web/lists/getByTitle('DocsNodePinnedLocations')/Fields/getbytitle('Modified By')",
					headers: {
						'Authorization': 'Bearer ' + tokenResponse.accessToken,
						'Accept': 'application/json; odata=verbose',
						'Content-Type': 'application/json; odata=verbose',
						'IF-MATCH': '*',
						'X-Http-Method': 'PATCH'
					},

					body: JSON.stringify({
						'__metadata': {
							// Type that you are modifying.
							'type': 'SP.FieldUser'
						},
						'ReadOnlyField': false
					}),
				};


				context.log(options);
				request(options, function (error, res, body) {
					// context.log(error);
					// context.log(body);
					// context.res = {
					//     body: body || ''
					// };
					// context.done();

					var itemProperties = JSON.stringify
						({
							__metadata:
							{
								type: listItemEntityType
							},
							'EditorId': 12,
							'Title': "Tesing12345",
							'_CreatorStringId': "12"
						});


					var optUser = {
						method: "POST",
						async: false,
						//uri: siteUrl + "/_api/web/lists/getbytitle('"+documentLibrary+"')/getItemByStringId('"+itemID+"')",
						uri: "https://docsnode.sharepoint.com/_api/web/lists/getbytitle('DocsNodePinnedLocations')/items/getById(28)",
						//	uri: "https://docsnode.sharepoint.com/sites/KnutSiteColl/_api/web/lists/getbytitle('AuthorUpdateList')/items/getById(1)",
						body: itemProperties,
						headers: {
							'Authorization': 'Bearer ' + tokenResponse.accessToken,
							'Accept': 'application/json; odata=verbose',
							'Content-Type': 'application/json; odata=verbose',
							'X-RequestDigest': reqDigest,
							'IF-MATCH': '*',
							'X-Http-Method': 'PATCH'


						}
					};
					request(optUser, function (error, res, body) {
						var optionsSetTrue = {
							method: "POST",
							uri: "https://docsnode.sharepoint.com/_api/web/lists/getByTitle('DocsNodePinnedLocations')/Fields/getbytitle('Created By')",
							headers: {
								'Authorization': 'Bearer ' + tokenResponse.accessToken,
								'Accept': 'application/json; odata=verbose',
								'Content-Type': 'application/json; odata=verbose',
								'IF-MATCH': '*',
								'X-Http-Method': 'PATCH'
							},

							body: JSON.stringify({
								'__metadata': {
									// Type that you are modifying.
									'type': 'SP.FieldUser'
								},
								'ReadOnlyField': false
							}),
						};
						request(optionsSetTrue, function (error, res, body) {
							var success = { success: "Creator updated" };
							resolve(success);
						});
					});

				});

				// return tokenResponse.accessToken;
			});
	});

}