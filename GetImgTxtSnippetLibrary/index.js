// var request = require("request");
const axios = require("axios");

var appTokenHelper = require("../lib/authTokenHelper");
var helpers = require("../helper/commonHelpers");


module.exports = async function (context, req) {

    try {
        var authorityHostUrl = 'https://login.microsoftonline.com';
        var tenant = ''; //'docsnode.com';
        var resource = '';
        var folderPath = '';
        var title = '';

        if (req.body && req.body.tenant && req.body.SPOUrl) {
            resource = req.body.SPOUrl;
            tenant = req.body.tenant;
            folderPath = req.body.FolderPath;
            title = req.body.Title;
        }
        const userToken = req.headers.authorization || null;
        let accesstoken = null;
        if (!userToken) {
            throw new HTTPError(401, 'Unauthorized. Missing authorization header.')
        }
        else {
            accesstoken = userToken;
        }

        var optionsImagesiteURL = {
            method: "GET",
            uri: resource + "/sites/DocsNodeAdmin/_api/web/lists/getbytitle('DocsNodeConfiguration')/items?$filter=ConfigAssestTitle eq '" + title + "'",
            headers: {
                'Authorization': accesstoken,
                'Accept': 'application/json; odata=verbose',
                'Content-Type': 'application/json; odata=verbose'
            },
        };


        const response = await axios.get(optionsImagesiteURL.uri, {
            headers: optionsImagesiteURL.headers
        });

        const formatJSON = response.data;

        var configlistGuid = formatJSON.d.results[0].ConfigSourceListGUID;
        var configSiteUrl = formatJSON.d.results[0].ConfigSourceListPath;
        var configListName = formatJSON.d.results[0].ConfigSourceList;
        var configSourceList = formatJSON.d.results[0].ConfigSourceDisplayListName;

        var configListRelUrl = configSiteUrl + "/" + configSourceList + "/" + folderPath;

      //  if (configListName === "Sample Org Assets") {
            var options = {
                method: "GET",
                uri: resource + "/" + configSiteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + configListRelUrl + "')?$expand=Files,Folders",
                headers: {
                    'Authorization': accesstoken,
                    'Accept': 'application/json; odata=verbose',
                    'Content-Type': 'application/json; odata=verbose'
                }
            };

            const assetResponse = await axios.get(options.uri, {
                headers: options.headers
            });

            context.res = {
                body: assetResponse ? helpers.safeStringify(assetResponse.data.d) : ''
            };
            context.done();

      //  }
       // else if (configListName === "DocsNodeText") {

            // logic for folder structure for SP list

         //   var options = {
           //     method: "GET",
             //   uri: resource + "/" + configSiteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + configListRelUrl + "')?$expand=Files,Folders",
              //  headers: {
                //    'Authorization': accesstoken,
                  //  'Accept': 'application/json; odata=verbose',
                   // 'Content-Type': 'application/json; odata=verbose'
                //}
            //};
            //const assetResponse = await axios.get(options.uri, {
              //  headers: options.headers
            //});

            //context.res = {
              //  body: assetResponse ? helpers.safeStringify(assetResponse.data.d) : ''
            //};
            //context.done();
        //}


    }
    catch (error) {
        // If the promise rejects, an error will be thrown and caught here
        context.log(error);
        context.done();
    }
};


// safely handles circular references

