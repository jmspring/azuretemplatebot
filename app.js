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
      ResourceManagementClient = require('azure-arm-resource').ResourceManagementClient,
      SubscriptionClient = require('azure-arm-resource').SubscriptionClient;
      
var azureAccountFields = {
    'subscriptionId': { 'text': 'Subscription Id', 'type': 'uuid' },
    'tenantId': { 'text': 'Tenant Id', 'type': 'uuid' },
    'clientId': { 'text': 'Client Id', 'type': 'uuid' },
    'clientSecret': { 'text': 'Client Secret', 'type': 'string' }
};

var templateOptions = {
    'generate': {
        'text': 'Select and enter information for a template deployment.',
        'template': '/generate_template',
        'requiresGenerate': false,
        'requiresResourceGroup': false
    },
    'show': {
        'text': 'Show the template generated.',
        'template': '/show_template',
        'requiresGenerate': true,
        'requiresResourceGroup': false,
    },
    'deploy': {
        'text': 'Deploy a generated template.',
        'template': '/deploy_template',
        'requiresGenerate': true,
        'requiresResourceGroup': true,
        'requiresName': true
    },
    'verify': {
        'text': 'Verify a generated template.',
        'template': '/verify_template',
        'requiresGenerate': true,
        'requiresResourceGroup': true,
        'requiresName': true
    },
    'resource_group': {
        'text': 'Specify or create a resource group for template operations.',
        'template': '/resource_group',
        'requiresGenerate': false,
        'requiresResourceGroup': false
    },
    'name': {
        'text': 'Specify name for deployment.  Required for deploy and verify.',
        'template': '/deployment_name',
        'requireGenerate': false,
        'requireResourceGroup': false,
        'requireName': false
    },
    'clear': {
        'text': 'Clear information on currently generated template.',
        'template': '/clear_template',
        'requiresGenerate': true,
        'requiresResourceGroup': false
    },
    'quit': {
        'text': 'Exit the bot.',
        'template': '/quit',
        'requiresGenerate': false,
        'requiresResourceGroup': false
    }
};

var resourceGroupOptions = {
    'create': {
        'text': 'Create a new resource group.',
        'template': '/create_resource_group',
        'promptForGroup': true
    },
    'delete': {
        'text': 'Delete existing resource group.',
        'template': '/delete_resource_group',
        'promptForGroup': false
    },
    'set': {
        'text': 'Set active resource group.',
        'template': '/set_resource_group',
        'promptForGroup': true
    },
    'clear': {
        'text': 'Clear active resource group.',
        'template': '/clear_resource_group',
        'promptForGroup': false
    },
    'show': {
        'text': 'Show current resource group.',
        'template': '/show_resource_group',
        'promptForGroup': false
    },
    'quit': {
        'text': 'Exit the bot.',
        'template': '/quit',
        'promptForGroup': false
    },
    'exit': {
        'text': 'Exit Resource Group menu.',
        'template': '/exit',
        'promptForGroup': false
    }
};

// Azure template files located within the template repository
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

function get_resource_client(account) {
    var credentials = new msRestAzure.ApplicationTokenCredentials(
                            account['clientId'],
                            account['tenantId'],
                            account['clientSecret']);
    var resourceClient = new ResourceManagementClient(
                            credentials,
                            account['subscriptionId']);
    return resourceClient;
}

function get_subscription_client(account) {
    var credentials = new msRestAzure.ApplicationTokenCredentials(
                            account['clientId'],
                            account['tenantId'],
                            account['clientSecret']);
    var subscriptionClient = new SubscriptionClient(credentials);
    return subscriptionClient;
}

// Dialogs
function isQuitting(session, results) {
    var response = results.response.toString().trim();
    return (response.toLowerCase() === 'quit');
}

function quit(session, results) {
    // TODO -- clean up userData
    session.userData.templateInfo = {};
    session.userData.azureAccount = {};
    session.userData.resourceGroup = {};
    session.userData['loggedIn'] = false;
    session.send("Come back anytime!");
    session.endConversation();
}

bot.dialog('/', [
    function(session) {
        if('loggedIn' in session.userData && session.userData['loggedIn'] == true) {
            session.beginDialog('/template_actions');
        } else {
            session.beginDialog('/login');
        }
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            session.replaceDialog('/');
        } else {
            quit(session, results);
        }
    }
]);

bot.dialog('/login', [
    function(session) {
        session.userData['loggedIn'] = false;
        session.userData.azureAccount = {};
        
        session.send("Welcome.  Please enter your Azure credentials.  You can stop at anytime by typing 'quit'");
        session.beginDialog('/enter_login_info');
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
            session.endDialog({response: 'quit'});
        }        
    },
    function(session, results) {
        var correct = results.response;
        if(correct) {
            // to verify login, attempt a simple resource manager operation
            var resourceClient = get_resource_client(session.userData.azureAccount);
            resourceClient.resources.list(function (err, result, request, response) {
                if(err) {
                    session.send("The credentials specified seem incorrect.  Please try again.")
                } else {
                    session.userData['loggedIn'] = true;
                } 
                session.endDialog();
            });
        } else {
            session.endDialog();
        }
    }
]);

bot.dialog('/enter_login_info', [
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
            session.replaceDialog('/enter_login_info');
        } else {
            session.endDialog({response: 'quit'});
            
        }
    }
]);

// have we properly generated the template?
// this currently relies on the template input setting the 'generated' flag.
// TODO - verify all required fields present and have values
function template_generated(templateInfo) {
    if(!templateInfo || 
            !('generated' in templateInfo) ||
            (templateInfo['generated'] != true)) {
        return false;
    } else {
        return true;
    }
}

function template_generate_parameters(templateInfo) {
    var parameters = JSON.parse(JSON.stringify(templateInfo['parameters']));
    var fields = Object.keys(parameters['parameters']);
    for(i = 0; i < fields.length; i++) {
        parameters['parameters'][fields[i]]['value'] = templateInfo[fields[i]];
    }
    return parameters;
}

function get_resource_group(session) {
    var resourceGroup = null;
    if(('resourceGroup' in session.userData) && session.userData.resourceGroup["name"]) {
        resourceGroup = session.userData.resourceGroup["name"];
    }
    
    return resourceGroup;
}

function requires_deployment_name(action) {
    if(('requiresName' in templateOptions[action]) && templateOptions[action]['requiresName']) {
        return true;
    }
    return false;
}

function get_deployment_name(session) {
    if(('templateInfo' in session.userData) && ('deploymentName' in session.userData.templateInfo)) {
        return session.userData.templateInfo['deploymentName'];
    }
    return null;
}

bot.dialog('/template_actions', [
    function(session) {
        var fields = Object.keys(templateOptions);
        for(var i = 0; i < fields.length; i++) {
            var field = fields[i];
            session.send(field + " --> " + templateOptions[field].text);
        }
        builder.Prompts.choice(session, "What would you like to do?", templateOptions);
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            var action = results.response.entity;
            var dialog = templateOptions[action];
            var resourceGroup = get_resource_group(session);
            var deploymentName = get_deployment_name(session);
            if(templateOptions[action]['requiresGenerate'] && !template_generated(session.userData.templateInfo)) {
                session.send("The action '" + action + "' requires that the template be generated first.");
                session.replaceDialog('/template_actions');
            } else {
                if(templateOptions[action]['requiresResourceGroup'] && !resourceGroup) {
                    session.send("The action '" + action + "' requires the resource group be specified.");
                    session.replaceDialog('/template_actions');
                } else {
                    if(requires_deployment_name(action) && !deploymentName) {
                        session.send("The action '" + action + "' requires a deployment name be specified.");
                        session.replaceDialog('/template_actions');
                    } else {
                        session.beginDialog(templateOptions[action].template);
                    }
                }
            }
        } else {
            session.endDialog({response: 'quit'});
        }   
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            session.endDialog();
        } else {
            session.endDialog({response: 'quit'});
        }
    }
]);

bot.dialog('/generate_template', [
    function(session) {
        session.userData.templateInfo = {};

        builder.Prompts.text(session, "What is the URL for the template repository?");
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            if(!validator.isURL(results.response)) {
                session.send("Please enter a valid URL.");
                session.replaceDialog('/generate_template');
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
                                session.beginDialog('/generate_template_parameters');
                            } else {
                                session.userData.templateInfo = {};
                                session.send("Unable to retrieve azuredeployment.parameters.json.  Try again.");
                                session.endDialog();
                            }
                        });
                    } else {
                        session.userData.templateInfo = {};
                        session.send("Unable to retrieve azuredeployment.json.  Try again.");
                        session.endDialog();
                    }
                });
            }
        } else {
            session.endDialog({response: 'quit'});
        }
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            session.send("Template values chosen:");
            var parameters = session.userData.templateInfo['parameters']
            var fields = Object.keys(parameters['parameters']);
            for(var i = 0; i < fields.length; i++) {
                var field = fields[i];
                session.send("    " + field + ": " + session.userData.templateInfo[field]);
            }
            builder.Prompts.confirm(session, "Are these values correct?"); 
        } else {
            session.endDialog({response: 'quit'});
        }
    },
    function(session, results) {
        var correct = results.response;
        if(correct) {
            session.userData.templateInfo['generated'] = true;
        } else {
            session.userData.templateInfo = {};
        }
        session.endDialog();
    }        
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

bot.dialog('/generate_template_parameters', [
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
            session.replaceDialog('/generate_template_parameters');
        } else {
            session.endDialog({response: 'quit'});
        }
    }
]);

bot.dialog('/show_template', [
    function(session) {
        // TODO -- actually show the generated JSON.  For now, just showing selected values.
        session.send("Template values chosen:");
        var parameters = session.userData.templateInfo['parameters']
        var fields = Object.keys(parameters['parameters']);
        for(var i = 0; i < fields.length; i++) {
            var field = fields[i];
            session.send("    " + field + ": " + session.userData.templateInfo[field]);
        }
        session.endDialog();    
    }
]);

function deploy_template(accountInfo, resourceGroup, deploymentName, template, templateParameters, callback) {
    var resourceClient = get_resource_client(accountInfo);
    var parameters = {
        'properties': {
            'template': template,
            'parameters': templateParameters,
            'mode': 'Incremental'
        }
    };
    resourceClient.deployments.createOrUpdate(resourceGroup, deploymentName, parameters, callback);
}

bot.dialog('/deploy_template', [
    function(session) {
        var template = session.userData.templateInfo['template'];
        var parameters = template_generate_parameters(session.userData.templateInfo);
        var resourceGroup = session.userData.resourceGroup['name'];
        var deploymentName = get_deployment_name(session);
        deploy_template(session.userData.azureAccount, 
                        resourceGroup,
                        deploymentName,
                        template,
                        parameters['parameters'],
                        function(err, result) {
            if(err) {
                session.send("Error deploying template.  Error: " + err);
            } else {
                session.send("Template deployed.");
            }
            session.endDialog();
        });
    }
]);

function verify_template(accountInfo, resourceGroup, deploymentName, template, templateParameters, callback) {
    var resourceClient = get_resource_client(accountInfo);
    var parameters = {
        'properties': {
            'template': template,
            'parameters': templateParameters,
            'mode': 'Incremental'
        }
    };
    resourceClient.deployments.validate(resourceGroup, deploymentName, parameters, callback);
}

bot.dialog('/verify_template', [
    function(session) {
        var template = session.userData.templateInfo['template'];
        var parameters = template_generate_parameters(session.userData.templateInfo);
        var resourceGroup = session.userData.resourceGroup['name'];
        var deploymentName = get_deployment_name(session);
        verify_template(session.userData.azureAccount, 
                        resourceGroup,
                        deploymentName,
                        template,
                        parameters,
                        function(err, result) {
            if(err) {
                session.send("Error validating template.  Error: " + err);
            } else {
                session.send("Template valid.");
            }
            session.endDialog();
        });
    }
]);

bot.dialog('/clear_template', [
    function(session) {
        session.userData.templateInfo = {};
        session.endDialog();
    }
]);

bot.dialog('/deployment_name', [
    function(session) {
        builder.Prompts.text(session, "Enter name for deployment.");
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            session.userData.templateInfo['deploymentName'] = results.response;
            session.endDialog();
        } else {
            session.endDialog({response: 'quit'});
        }
    }
]);

function check_for_resource_group(accountInfo, resourceGroup, callback) {
    var resourceClient = get_resource_client(accountInfo);
    resourceClient.resourceGroups.checkExistence(resourceGroup, callback);
}

function list_locations(accountInfo, callback) {
    var subscriptionClient = get_subscription_client(accountInfo);
    subscriptionClient.subscriptions.listLocations(accountInfo['subscriptionId'], callback);
}

function create_resource_group(accountInfo, resourceGroup, location, callback) {
    var resourceClient = get_resource_client(accountInfo);
    resourceClient.resourceGroups.createOrUpdate(resourceGroup, { 'location': location }, callback);    
}

function delete_resource_group(accountInfo, resourceGroup, callback) {
    var resourceClient = get_resource_client(accountInfo);
    resourceClient.resourceGroups.deleteMethod(resourceGroup, callback);    
}

bot.dialog('/resource_group', [
    function(session) {
        if(!('resourceGroup' in session.userData)) {
            session.userData['resourceGroup'] = {};
        }
        
        session.send("Resource group:");
        var fields = Object.keys(resourceGroupOptions);
        for(var i = 0; i < fields.length; i++) {
            var field = fields[i];
            session.send("    " + field + " --> " + resourceGroupOptions[field].text);
        }
        builder.Prompts.choice(session, "What would you like to do?", resourceGroupOptions);
    },
    function(session, results) {
        var action = results.response.entity;
        if(resourceGroupOptions[action]['promptForGroup'] == true) {
            session.userData.resourceGroup['action'] = action;
            builder.Prompts.text(session, "Please enter the resource group name.");
        } else {
            var dialog = resourceGroupOptions[action];
            session.beginDialog(resourceGroupOptions[action].template);
        }
    },
    function(session, results) {
        if(('action' in session.userData.resourceGroup) && (session.userData.resourceGroup['action'])){
            var action = session.userData.resourceGroup['action'];
            session.userData.resourceGroup['action'] = null;
            session.userData.resourceGroup['name'] = results.response;
            var dialog = resourceGroupOptions[action];
            session.beginDialog(resourceGroupOptions[action].template);
        } else {                            
            if(!isQuitting(session, results)) {
                if(results.response === 'exit') {
                    session.endDialog();
                } else {
                    session.replaceDialog('/resource_group');
                }
            } else {
                session.endDialog({response: 'quit'});
            }
        }
    },
    function(session, results) {
        if(!isQuitting(session, results)) {
            session.replaceDialog('/resource_group');
        } else {
            session.endDialog({response: 'quit'});
        }
    }
]);


bot.dialog('/create_resource_group', [
    function(session) {
        var resourceGroup = session.userData.resourceGroup['name'];
        check_for_resource_group(session.userData.azureAccount, 
                                 resourceGroup,
                                 function(err, result) {
            if(err) {
                session.userData.resourceGroup['name'] = null;
                session.send("An error occurred checking for resource group: " + resourceGroup);
                session.endDialog();
            } else {
                if(result) {
                    session.userData.resourceGroup['name'] = null;
                    session.send("Resource group " + resourceGroup + " already exists.");
                    session.endDialog();
                } else {
                    var locations = {};
                    list_locations(session.userData.azureAccount, function(err, result) {
                        if(err) {
                            session.send("An error occurred trying to retrieve locations.");
                            session.endDialog();
                        } else {
                            session.send("Locations for Resource Group:");
                            for(var i = 0; i < result.length; i++) {
                                session.send("  " + result[i].name + " --> " + result[i].displayName);
                                locations[result[i].name] = result[i].displayName;
                            }
                            builder.Prompts.choice(session, "Choose a location.", locations);
                        }
                    });
                }
            }                    
        });
    },
    function(session, results) {
        var resourceGroup = session.userData.resourceGroup['name'];
        session.userData.resourceGroup['name'] = null;
        var location = results.response.entity;

        create_resource_group(session.userData.azureAccount, resourceGroup, location, function(err, result) {
            if(err) {
                session.send("An error occurred attempting to create resource group: " + resourceGroup + ", at location: " + location);
            } else {
                session.send("Resource group created: " + resourceGroup);
            }
            session.endDialog();
        });
    }
]);

bot.dialog('/set_resource_group', [
    function(session) {
        var resourceGroup = session.userData.resourceGroup['name'];
        check_for_resource_group(session.userData.azureAccount, 
                                 resourceGroup,
                                 function(err, result) {
            if(err) {
                session.userData.resourceGroup['name'] = null;
                session.send("An error occurred checking for resource group: " + resourceGroup);
                session.endDialog();
            } else {
                if(!result) {
                    session.userData.resourceGroup['name'] = null;
                    session.send("Resource group " + resourceGroup + " not found.");
                }
                session.endDialog();
            }
        });
    }
]);

bot.dialog('/clear_resource_group', [
    function(session) {
        if(('name' in session.userData.resourceGroup) && session.userData.resourceGroup['name']) {
            session.userData.resourceGroup['name'] = null;
        }
        session.endDialog();
    }
]);

bot.dialog('/show_resource_group', [
    function(session) {
        if(('name' in session.userData.resourceGroup) && session.userData.resourceGroup['name']) {
            session.send("Current resource group set to: " + session.userData.resourceGroup['name']);
        } else {
            session.send("No resource group currently set.");
        }
        session.endDialog();
    }

]);

bot.dialog('/delete_resource_group', [
    function(session) {
        var resourceGroup = session.userData.resourceGroup['name'];
        if(!resourceGroup) {
            session.send("Must set resource group before deleteing.");
            session.endDialog();
        } else {
            check_for_resource_group(session.userData.azureAccount, 
                                     resourceGroup,
                                     function(err, result) {
                if(err) {
                    session.userData.resourceGroup['name'] = null;
                    session.send("An error occurred checking for resource group: " + resourceGroup);
                    session.send("Clearning resource group from selected resource group.");
                    session.endDialog();
                } else {
                    builder.Prompts.confirm(session, "Delete resource group: " + resourceGroup + "?  (The synchronous operation can take some time...)");
                }
            });
        }
    },
    function(session, results) {
        var confirm = results.response;
        if(confirm) {
            var resourceGroup = session.userData.resourceGroup['name'];
            session.userData.resourceGroup['name'] = null;
            delete_resource_group(session.userData.azureAccount,
                                  resourceGroup, 
                                  function(err, result) {
                if(err) {
                    session.send("An error occurred deleting the resource group: " + resourceGroup);
                } else {
                    session.send("Resource group deleted: " + resourceGroup);
                }
                session.endDialog();
            });
        } else {
            session.endDialog();
        }
    }
]);

bot.dialog('/quit', [
    function(session) {
        session.endDialog({response: 'quit'});
    }
]);

bot.dialog('/exit', [
    function(session) {
        session.endDialog({response: 'exit'});
    }
]);