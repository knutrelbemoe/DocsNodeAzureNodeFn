// var request = require("request");
const axios = require("axios");

var appTokenHelper = require("../lib/authTokenHelper");
var helpers = require("../helper/commonHelpers");


module.exports = async function (context, req) {

    try {
        var authorityHostUrl = 'https://login.microsoftonline.com';
        var tenant = ''; //'docsnode.com';
        var resource = '';
        var FolderPath = '';
        var title = '';
        let LibraryName = '';
        let TextSnippetContent = '';
        let TextSnippetFile = '';
        let TextSnippetDescription = '';
        let TargetSiteCollection = '';

        if (req.body && req.body.tenant && req.body.SPOUrl && req.body.LibraryName && req.body.TextSnippet) {
            resource = req.body.SPOUrl;
            tenant = req.body.tenant;
            TargetSiteCollection = req.body.TargetSite
            FolderPath = req.body.FolderPath;

            LibraryName = req.body.LibraryName;
            TextSnippetContent = req.body.TextSnippet;
            TextSnippetFile = req.body.SnippetTitle;
            TextSnippetDescription = req.body.SnippetDescription || null;
        }
        const userToken = req.headers.authorization || null;
        let accesstoken = null;
        if (!userToken) {

            const token = await appTokenHelper(resource, tenant);
            accesstoken = 'Bearer ' + token.accessToken;

            // throw new HTTPError(401, 'Unauthorized. Missing authorization header.')
        }
        else {
            accesstoken = userToken;
        }


        var buffer = Buffer.from(TextSnippetContent, "utf-8");

        var targetUrl = TargetSiteCollection + "/" + LibraryName + FolderPath;
        var url = resource + TargetSiteCollection + "/_api/Web/GetFolderByServerRelativeUrl(@target)/Files/add(overwrite=false, url='" + TextSnippetFile + ".txt')?$expand=ListItemAllFields&@target='" + targetUrl + "'";

        var optionsRequestData = {
            data: buffer,
            uri: url,
            headers: {
                'Authorization': accesstoken,
                'Accept': 'application/json; odata=verbose',
                'content-length': buffer.byteLength
            },
        };


        const response = await axios.post(optionsRequestData.uri, buffer, {
            headers: optionsRequestData.headers
        });


        if (response.data.d.ListItemAllFields.ID && TextSnippetDescription) {

            // const listUrl = resource + TargetSiteCollection + "/_api/web/getFileByServerRelativeUrl('" + response.data.d.ServerRelativeUrl + "')/ListItemAllFields";
            const listUrl = resource + TargetSiteCollection + "/_api/web/lists/GetByTitle('" + LibraryName + "')/items/getbyid(" + response.data.d.ListItemAllFields.ID + ")";
            const optionListMetadata = {
                "__metadata": { "type": "SP.Data." + LibraryName.charAt(0).toUpperCase() + LibraryName.split(" ").join("").slice(1) + "Item" },
                "SnippetDescription": TextSnippetDescription,
                "Title": TextSnippetFile

            }

            const assetResponse = await axios.post(listUrl, JSON.stringify(optionListMetadata), {
                headers: {
                    'Authorization': accesstoken,
                    'Accept': 'application/json; odata=verbose',
                    'Content-Type': 'application/json; odata=verbose',
                    "X-HTTP-Method": "MERGE",
                    "If-Match": "*"
                }
            });

            context.res = {
                body: assetResponse ? helpers.safeStringify({ message: 'File Upload Successful with description' }) : ''
            };


            context.done();

        }
        else {

            context.res = {
                body: helpers.safeStringify({ message: 'File Upload Successful' })
            };
            context.done();
        }


    }
    catch (error) {
        // If the promise rejects, an error will be thrown and caught here
        context.log(error);
        context.res = {
            body: helpers.safeStringify(error.response.data)
        };
        context.done();
    }
};


// safely handles circular references

