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
      validator = require('validator'),
      request = require('request'),
      msRestAzure = require('ms-rest-azure'),
      ResourceManagementClient = require('azure-arm-resource').ResourceManagementClient;
      
var azureAccountFields = {
    'subscriptionId': { 'text': 'Subscription Id', 'type': 'uuid' },
    'tenantId': { 'text': 'Tenant Id', 'type': 'uuid' },
    'clientId': { 'text': 'Client Id', 'type': 'uuid' },
    'clientSecret': { 'text': 'Client Secret', 'type': 'string' }
};

var templateOptions = [ "deploy", "check" ];
var azureTemplateName = "azuredeploy.json";
var azureParametersName = "azuredeploy.parameters.json";
      
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

function generate_url(url, file) {
    if(url.charAt(url.length - 1) != '/') {
        url = url + "/";
    }
    url = url + file;
    return url;
}

// Dialogs
function isQuitting(session, results) {
    var response = results.response.toString().trim();
    return (response.toLowerCase() === 'quit');
}

function quit(session, results) {
    session.send("Come back anytime!");
    session.userData.azureAccount = {};
    session.userData.templateInfo = {};
    session.endDialog();
}

bot.dialog('/', [
    function(session) {
        session.userData.azureAccount = {};
        session.userData.templateInfo = {};
        session.send("Welcome!  You can end the process at anytime by typing 'quit'");
        session.beginDialog('/login');
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
                session.beginDialog('/verify_login')
            } else {
                quit(session, results);
            }
        } else {
            quit(session, results);
        }
    }, 
    function(session, results) {
        if(!isQuitting(session, results)) {
            builder.Prompts.choice(session, "What would you like to do?", templateOptions);
        } else {
            quit(session, results);
        }   
    },
    function(session, results) {
        if(results.response) {
            if(results.response.entity === 'deploy') {
                session.beginDialog('/deploy_template');
            } else {
                session.send("Your choice was: " + results.response.entity);
            }
        } else {
            session.send("ok");
        }
    },
    function(session, results) {
        session.send("thanks!")
        session.endConversation();
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

bot.dialog('/verify_login', [
    function(session) {
        // to verify login, attempt a simple resource manager operation
        var credentials = new msRestAzure.ApplicationTokenCredentials(
                                session.userData.azureAccount['clientId'],
                                session.userData.azureAccount['tenantId'],
                                session.userData.azureAccount['clientSecret']);
        var resourceClient = new ResourceManagementClient(
                                credentials,
                                session.userData.azureAccount['subscriptionId']);
        
        resourceClient.resources.list(function (err, result, request, response) {
            if(err) {
                session.send("The credentials specified seem incorrect.  Please restart.")
                session.endConversation();
            } else {
                session.endDialog();
            } 
        });
    }
]);

bot.dialog('/deploy_template', [
    function(session) {
        builder.Prompts.text(session, "What is the URL for the template repository?");
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            if(!validator.isURL(results.response)) {
                session.send("Please enter a valid URL.");
                session.replaceDialog('/deploy_template');
            } else {
                // verify that the template file is present
                var templateUrl = generate_url(results.response, azureTemplateName);
                request(templateUrl, function (error, response, body) {
                    if (!error && response.statusCode == 200) {
                        session.userData.templateInfo["template"] = JSON.parse(body);
                         
                        // verify that the parameters file is present
                        var parametersUrl = generate_url(results.response, azureParametersName);                
                        request(parametersUrl, function(error, response, body) {
                            if (!error && response.statusCode == 200) {
                                session.userData.templateInfo["parameters"] = JSON.parse(body);
                                session.beginDialog('/deploy_template_parameters');
                            } else {
                                session.send("error retrieving parameters.  start over.");
                                session.endDialog({response: 'quit'});
                            }
                        });
                    } else {
                        session.send("error retrieving parameters.  start over.");
                        session.endDialog({response: 'quit'});
                    }
                });
            }
        }
    },
]);

function prompt_for_template_value(session, key) {
    var template = session.userData.templateInfo["template"];
    var parameters = session.userData.templateInfo["parameters"];
    
    if("allowedValues" in template["parameters"][key]) {
        var choices = template["parameters"][key]["allowedValues"];
        builder.Prompts.choice(session, "Enter value for " + key + ".", choices);
    } else {
        builder.Prompts.text(session, "Enter value for " + key + ".");
    }
}

bot.dialog('/deploy_template_parameters', [
    function(session) {
        var fields = Object.keys(session.userData.templateInfo['parameters']['parameters']);
        session.userData.templateInfo['currentKey'] = null;
        for(var i = 0; i < fields.length; i++) {
            var field = fields[i];
            if(!(field in session.userData.templateInfo) || !session.userData.templateInfo[field]) {
                session.userData.templateInfo['currentKey'] = field;
                prompt_for_template_value(session, field);
                break;
            }
        }
        if(session.userData.templateInfo['currentKey'] == null) {
            session.endDialog();
        }
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            var template = session.userData.templateInfo['template'];
            var parameters = session.userData.templateInfo['parameters']
            var fields = Object.keys(parameters['parameters']);
            var key = session.userData.templateInfo['currentKey'];
            if(!("allowedValues" in template["parameters"][key])) {
                var value = results.response;
                // string should not be empty
                if(value) {
                    session.userData.templateInfo[key] = value;
                } else {
                    session.send("You must enter a valid string for " + key + ".");
                }
            } else {
                session.userData.templateInfo[key] = results.response.entity;
            }
            session.replaceDialog('/deploy_template_parameters');
        } else {
            session.endDialog({response: 'quit'});
        }
    }
]);


