// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import path from 'path';
import restify from 'restify';
import { adapter, EchoBot } from './bot';
import tabs from './tabs';
import MessageExtension from './message-extension';

// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { ActivityTypes } from 'botbuilder';
let embedToken = require(__dirname + '/embedConfigService.js');
const utils = require(__dirname + "/utils.js");

//Create HTTP server.
const server = restify.createServer({
    formatters: {
        'text/html': function (req, res, body) {
            return body;
        },
    },
});
// // Prepare server for Bootstrap, jQuery and PowerBI files

server.get('/js/bootstrap', restify.plugins.serveStatic({
    directory: 'node_modules/bootstrap/dist/js/',
    file: 'bootstrap.min.js'
  }));

  server.get('/js/powerbi', restify.plugins.serveStatic({
    directory: 'node_modules/powerbi-client/dist/',
    file: 'powerbi.min.js'
  }));
  server.get('/js/jquery', restify.plugins.serveStatic({
    directory: 'node_modules/jquery/dist/',
    file: 'jquery.min.js'
  }));

  server.use(restify.plugins.jsonBodyParser());
 server.use(restify.plugins.urlEncodedBodyParser( {extended: true}));


server.get('/getEmbedToken', async function (req, res) {

    // // Validate whether all the required configurations are provided in config.json
    // configCheckResult = utils.validateConfig();
    // if (configCheckResult) {
    //     return res.status(400).send({
    //         "error": configCheckResult
    //     });
    // }
    // Get the details like Embed URL, Access token and Expiry
    let result = await embedToken.getEmbedInfo();

    // result.status specified the statusCode that will be sent along with the result object
    res.status(result.status)
    res.send(result);
});

// Read botFilePath and botFileSecret from .env file.
const ENV_FILE = path.join(__dirname, '.env');
require('dotenv').config({ path: ENV_FILE });



server.get(
    '/*',
    restify.plugins.serveStatic({
        directory: __dirname + '/static',
    })
);

server.listen(process.env.port || process.env.PORT || 3333, function () {
    console.log(`\n${server.name} listening to ${server.url}`);
});

// Adding tabs to our app. This will setup routes to various views
tabs(server);

// Adding a bot to our app
const bot = new EchoBot();

// Adding a messaging extension to our app
const messageExtension = new MessageExtension();

// Listen for incoming requests.
server.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        if (context.activity.type === ActivityTypes.Invoke)
            await messageExtension.run(context);
        else await bot.run(context);
    });
});
