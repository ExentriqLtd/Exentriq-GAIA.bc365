module.exports = function (RED) {
    "use strict";
    const axios = require('axios');
    const FormData = require('form-data');
    const soap = require('soap');

    function businessCentralConfigNode(config) {
        RED.nodes.createNode(this, config);
        var node = this;
        node.getConfig = () => {
            return config
        }
        // console.log("cbusinessCentralNode:: config", config)
    }

    RED.nodes.registerType("BC365Config", businessCentralConfigNode);


    function businessCentralNode(config) {
        RED.nodes.createNode(this, config);
        const node = this;
        const nodeSetting = RED.nodes.getNode(config.setting)
        const configSetting = nodeSetting.getConfig();


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
                var serverHost = config.server;
                const client = await getClientSoap(accessToken, urlMethod, serverHost);
                client.setSecurity(new soap.BearerSecurity(accessToken));
                if (!client) throw 'Erro Get Client';
                if (!client[paramNameJustify]) throw 'methods name not found';
                client[paramNameJustify](args, function (err, result) {
                    if (err) {
                        node.error(err)
                    } else {
                        let lastResponse =[];
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
                console.log('Error Oauth2 soap request');
                cb(error)
            }
        }

        /**
         * get client soap
         * @param {*} accessToken 
         * @param {*} SOAPUrl 
         * @param {*} serverHost
         * @returns 
         */
        async function getClientSoap(accessToken, SOAPUrl, serverHost) {
            let client = null;
            try {
                const options = {
                    wsdl_headers: {
                        "Authorization": "Bearer " + accessToken
                    }
                }
                const urlClient = `${serverHost}/wsdlDynamic?SOAPUrl=${SOAPUrl}&access_token=${accessToken}`;
                client = await soap.createClientAsync(urlClient, options)
            } catch (error) {
                console.error('error::WSDL::', error)
            }

            return client;
        }
        
        /**
        * event "input", when the flux start by inject/httpRequest
        */
        node.on('input', async function (msg, nodeSend, nodeDone) {
            if (!msg.payload) {
                msg.payload = 'payload empty';
                nodeSend(msg);
                return msg;
            }
            try {
                const methodName = configSetting.servicesMethods.split(';')[0];
                const urlMethod = configSetting.servicesDropDown.split(';')[1];
                executeSoap(configSetting, urlMethod, methodName, msg.payload, (res) => { //pass the corrected Payload in relation with the output in front-end;
                    msg.payload = res;
                    nodeSend(msg);
                    return msg;
                }); 
                msg.payload = "ok";
                return msg;


            } catch (error) {
                console.log('error:::::::on:input', error);
            }
        });
    }
    RED.nodes.registerType("BC365Connection", businessCentralNode);
}