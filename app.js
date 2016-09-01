// azuretempaltebot
// An interactive bot for deploying Azure Resource Manager templates from
// github or similar repositories.
//
// Requirements:
// - subscription id
// - tenant id
// - client id
// - client secret
// - url to the repository containing the template

const builder = require('botbuilder'),
      restify = require('restify'),
      arm = require('azure-arm-resource');
      
var azureAccountFields = {
    'subscriptionId': { 'text': 'Subscription Id', 'type': 'uuid' },
    'tenantId': { 'text': 'Tenant Id', 'type': 'uuid' },
    'clientId': { 'text': 'Client Id', 'type': 'uuid' },
    'clientSecret': { 'text': 'Client Secret', 'type': 'string' }
}
      
// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
    console.log('%s listening to %s', server.name, server.url);
});

// Create chat bot
var connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});
var bot = new builder.UniversalBot(connector);
server.post('/api/messages', connector.listen());
server.get(/.*/, restify.serveStatic({
    'directory': __dirname,
    'default': 'index.html'
}));

function validate_uuid(str) {
  return /[0-9a-f]{22}|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i.test(str);
}

// Dialogs
function isQuitting(session, results) {
    var response = results.response.toString().trim();
    return (response.toLowerCase() === 'quit');
}

function quit(session, results) {
    session.send("Come back anytime!")
    session.userData.azureAccount = {}
    session.userData.templateInfo = {}
    session.endDialog();
}

bot.dialog('/', [
    function(session) {
        session.userData.azureAccount = {}
        session.userData.templateInfo = {}
        session.send("Welcome!  You can end the process at anytime by typing 'quit'");
        session.beginDialog('/login')
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            session.send("Information:")
            var fields = Object.keys(azureAccountFields);
            for(var i = 0; i < fields.length; i++) {
                var field = fields[i];
                session.send("    " + azureAccountFields[field]['text'] + ": " + session.userData.azureAccount[field]);
            }
            builder.Prompts.confirm(session, "Are these values correct?"); 
        } else {
            quit(session, results);
        }
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            var correct = results.response;
            if(correct) {
                session.send("Great!  Let's get started.");
                session.endDialog();
            } else {
                quit(session, results);
            }
        } else {
            quit(session, results);
        }
    }
]);

bot.dialog('/login', [
    function(session) {
        var fields = Object.keys(azureAccountFields);
        session.userData.azureAccount['currentKey'] = null;
        for(var i = 0; i < fields.length; i++) {
            var field = fields[i];
            if(!(field in session.userData.azureAccount) || !session.userData.azureAccount[field]) {
                session.userData.azureAccount['currentKey'] = field;
                builder.Prompts.text(session, "Please enter your " + azureAccountFields[field]['text'] + ".");
                break;
            }
        }
        if(session.userData.azureAccount['currentKey'] == null) {
            session.endDialog();
        }
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            var fields = Object.keys(azureAccountFields);
            var key = session.userData.azureAccount['currentKey'];
            var value = results.response;
            if(azureAccountFields[key]['type'] === 'uuid') {
                // validate uuid
                if(validate_uuid(value)) {
                    session.userData.azureAccount[key] = value;
                } else {
                    session.send("You entered an invalid ID for " + azureAccountFields[key]['text'] + ".");
                }
            } else {
                // string should not be empty
                if(value) {
                    session.userData.azureAccount[key] = value;
                } else {
                    session.send("You must enter a valid string for " + azureAccountFields[key]['text'] + ".");
                }
            }
            session.replaceDialog('/login');
        } else {
            session.endDialog({response: 'quit'});
        }
    }
]);