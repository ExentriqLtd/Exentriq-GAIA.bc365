module.exports = function (RED) {
    "use strict";
    const axios = require('axios');
    const FormData = require('form-data');
    const express = require('express');
    const cors = require('cors');
    const app = express();
    const port = 20220;//3000
    const soap = require('soap');
  

    ///istanza1.nodered.it:1880 > nodered -> server url istanza1.nodered.it:20220
    ///istanza.nodered.it:20220 -> express app for wrapping request

    app.use(cors({
        origin: '*'
    }));

    app.listen(port, () => {
        console.log(`Server listening at the port: ${port}`);
    });

    function businessCentralNode(config) {
        RED.nodes.createNode(this, config);
        const node = this;
        /**
         * get Oauth2 and the xml of web Services to pass to mini-Server
         */
        app.get('/wrapperoauth2', (req, res) => {
            getOauth2(config)
                .then(async (response) => {
                    const url = config.baseUrl+config.tenant+"/"+config.environment+"/ODataV4/Company(\'"+config.company+"\')/"+config.append;
                    if (!response?.data?.access_token) return res.json({ status: 401, reason: 'access token not exist' });
                    const responseSoap = await getRequestDynamic(url, response?.data?.access_token);
                    res.json(responseSoap.data);
                })
                .catch((e) => {
                    console.log('error:::wrapperoauth2::', e)
                });
        });

        /**
         *  Oatuh2 -> wsdl of xml
         */
        app.get('/wsdlDynamic', (req, res) => {

            getOauth2(config)
                .then(async (response) => {
                    const url = req.query.SOAPUrl;
                    if (!response?.data?.access_token) return res.json({ status: 401, reason: 'access token not exist' });
                    const responseMethod = await getRequestDynamic(url, response?.data?.access_token);
                    res.set('Content-Type', 'text/xml');
                    res.send(responseMethod.data)
                })
                .catch((error) => {
                    console.error(error);
                    res.json({ status: 500, error });
                })
        })

        /**
         * pass to the mini server the Methods and the respective parameters of the web services list
         */
        app.get('/methods', (req, res) => {
            getOauth2(config)
                .then(async (response) => {
                    const url = req.query.SOAPUrl;
                    const methodName = req.query.methodName;

                    if (!url || !methodName) return console.warn('missed url / methodName')
                    if (!response?.data?.access_token) return res.json({ status: 401, reason: 'access token not exist' });
                    try {
                        //get client and describe and format array, response to front-end 
                        const client = await getClientSoap(response.data.access_token, url,config);
                        var methodNameJustify = methodName.replace(/\. |\s/g, "_");
                        //regex for names which includes '.' or ' ' -> '_'
                        let soapURLName = url.includes('/Page/') ? methodNameJustify + '_Service' : methodNameJustify;
                        const clientDescribe = client.describe();
                        const methodsList = clientDescribe[soapURLName][methodNameJustify + '_Port']
                        console.log('methodsList:::::', methodsList);
                        const functionNames = Object.keys(methodsList);
                        const parameterMapping = {};

                        for (const functionName in methodsList) {
                            if (Object.hasOwnProperty.call(methodsList, functionName)) {
                                const functionParams = methodsList[functionName].input;
                                parameterMapping[functionName] = functionParams;
                            }
                        }

                        console.log('parameterMapping', parameterMapping);
                        res.json(parameterMapping);
                    } catch (error) {
                        console.error(error);
                        res.json({ status: 500, error });
                    }
                })
                .catch((e) => {
                    console.log('error:::methods', e)
                });
        });

        /**
         * event "input", when the flux start by inject/httpRequest
         */
        node.on('input', async function (msg, nodeSend, nodeDone) {
            if (!msg.payload) {
                msg.payload = 'non ci sono informazione nel payload';
                nodeSend(msg);
                return msg;
            }
            try {
                const methodName = config.servicesMethods.split(';')[0];
                const urlMethod = config.servicesDropDown.split(';')[1];
                executeSoap(config, urlMethod, methodName, msg.payload, (res) => {//pass the corrected Payload in relation with the output in front-end;
                    msg.payload = res;
                    console.log('msg.payload:::::::', msg.payload);
                    nodeSend(msg);
                    //nodeDone(msg);
                    return msg;
                }); 

            } catch (error) {
                console.log('error:::::::on:input', error);
            }
        });
    }

    /**
     * get client soap
     * @param {*} accessToken 
     * @param {*} SOAPUrl 
     * @returns 
     */
    async function getClientSoap(accessToken, SOAPUrl,config) {
        let client = null;
        try {
            const options = {
                wsdl_headers: {
                    "Authorization": "Bearer " + accessToken
                }
            }
            const urlClient = config.server+'/wsdlDynamic?SOAPUrl=' + SOAPUrl
            client = await soap.createClientAsync(urlClient, options)
        } catch (error) {
            console.error('error::', error)
        }

        return client;
    }

    /**
     * 
     * Function that, after inserting the parameters, allows me to make the SOAP POST of the request obtaining the result in the client.
     * @param {*} config 
     * @param {*} urlMethod 
     * @param {*} methodName 
     * @param {*} args 
     * @param {*} cb
     * @returns 
     */
    async function executeSoap(config, urlMethod, paramName, args, cb) {
        try {
            const response = await getOauth2(config);
            if (!response?.data?.access_token) return res.json({ status: 401, reason: 'access token not exist' });

            const accessToken = response.data.access_token;
            const metName = config.servicesDropDown.split(';')[0];
            var methodNameJustify = metName.replace(/\. |\s/g, "_");
            var paramNameJustify = paramName.replace(/\. |\s/g, "_");

            const client = await getClientSoap(accessToken, urlMethod,config);
            client.setSecurity(new soap.BearerSecurity(accessToken));

            if (!client) throw 'Errore nella Get Client';
            console.log('client:::', client);
            console.log('methodName_client:::', paramNameJustify);

            if (!client[paramNameJustify]) throw 'methods name not found';

            client[paramNameJustify](args, function (err, result) {
                if (err) {
                    console.error('ERRORCLIENT:::::', err);
                } else {
                    // Elabora il risultato della chiamata qui
                    console.log('result:::::::', result);

                    //console.log('lastRequest:::::', client.lastRequest);
                    //console.log('lastResponse:::::', client.lastResponse);
                    let lastResponse =[];
                    //if (!result[paramNameJustify+"_Result"]) throw 'The Request has no Results';
                    if(result[paramNameJustify+"_Result"]!=null){
                        lastResponse = result[paramNameJustify+"_Result"][methodNameJustify];
                        console.log('result:::::::', result[paramNameJustify+"_Result"][methodNameJustify]);
                    }else{
                        lastResponse.push('The Request has no Results');
                    }
            

                    cb(lastResponse);
                }
            });

        } catch (error) {
            console.log('Errore nell\'Oauth2 per la richiesta Soap');
            cb(error)
        }
    }

    /**
     * to get Oauth2
     * @param {*} config 
     * @returns 
     */
    const getOauth2 = async (config) => {
        const tokenURL = 'https://login.microsoftonline.com/' + config.tenant + '/oauth2/V2.0/token';
        const data = new FormData();
        data.append('client_id', config.clientID);
        data.append('client_secret', config.clientSecret);
        data.append('grant_type', config.grantType);
        data.append('scope', config.scope);
        data.append('tenant', config.tenant);
        data.append('code', config.code);
        data.append('redirect_uri', config.redirectUri);
        const configurationAxios = {
            method: 'post',
            maxBodyLength: Infinity,
            url: tokenURL,
            headers: {
                ...data.getHeaders()
            },
            data
        };
        return await axios.request(configurationAxios);
    }

    /**
     * To get Request/WSDL from the selected URL
     * @param {*} url 
     * @param {*} accessToken 
     * @returns 
     */
    const getRequestDynamic = async (url, accessToken) => {
        const Authorization = `Bearer ${accessToken}`
        const config = {
            method: 'get',
            maxBodyLength: Infinity,
            url,
            headers: {
                Authorization
            }
        };
        return await axios.request(config);
    }

    RED.nodes.registerType("business-central-connection", businessCentralNode);
}