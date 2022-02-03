// var request = require("request");
const axios = require("axios");

var appTokenHelper = require("../lib/authTokenHelper");
var helpers = require("../helper/commonHelpers");


module.exports = async function (context, req) {

    try {
        var authorityHostUrl = 'https://login.microsoftonline.com';
        var tenant = ''; //'docsnode.com';
        var resource = '';
        var ImagePath = '';
        var title = '';


        if (req.body && req.body.tenant && req.body.SPOUrl && req.body.ImagePath) {
            resource = req.body.SPOUrl;
            tenant = req.body.tenant;

            ImagePath = req.body.ImagePath;

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



        var targetUrl = ImagePath + "/OpenBinaryStream";

        var optionsRequestData = {
            uri: targetUrl,
            headers: {
                'Authorization': accesstoken,
                'Accept': 'application/json; odata=verbose',

            },

        };


        const response = await axios.get(optionsRequestData.uri, {
            headers: optionsRequestData.headers,
            responseType: 'arraybuffer',
        });


        if (response.data) {

            //var bytes = new Uint8Array(response.data);
            // var blob = new Blob([bytes], { type: "image/jpeg" });

            var readerResponse = response.data.toString('base64');

            context.res = {
                body: readerResponse ? helpers.safeStringify(readerResponse) : ''
            };


            context.done();

        }
        else {

            context.res = {
                body: helpers.safeStringify({ message: 'Failed in request' })
            };
            context.done();
        }


    }
    catch (error) {
        // If the promise rejects, an error will be thrown and caught here
        context.log(error);
        context.res = {
            body: helpers.safeStringify(error)
        };
        context.done();
    }
};

const readBlob = async (blob) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onend = reject;
        reader.onabort = reject;
        reader.onload = () => resolve(reader.result);
        reader.readAsDataURL(blob);
    });
};


// safely handles circular references

