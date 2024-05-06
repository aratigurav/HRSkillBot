// Import required packages
import * as restify from "restify";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  ConfigurationServiceClientCredentialFactory,
} from "botbuilder";

import {
  allowedCallersClaimsValidator,
  AuthenticationConfiguration,
  AuthenticationConstants
} from "botframework-connector";

// This bot's main dialog.
import app from "./app";
import config from "./config";

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3979, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// Expose the manifest
server.get('/manifest/*', restify.plugins.serveStatic({ directory: './manifest', appendRequestPath: false }));

const allowedCallers: string[] = ['http://localhost:3978','68508543-951e-443d-8200-a2672f3d5193'];//(process.env.BOT_ALLOWED_CALLERS || '').split(',').filter((val) => val); //['http://localhost:3978','8183e562-4882-4d3f-8cf8-bac799a08b99'];
console.log('allowed caller ' + allowedCallers[0], allowedCallers[1]);
//const allowedCallers = (process.env.AllowedCallers || '').split(',').filter((val) => val) || [];

var claimsValidators;
if (allowedCallers.length > 0) {
   claimsValidators = allowedCallersClaimsValidator(allowedCallers);
  // Continue with further processing using claimsValidators
} else {
  // Handle the case when allowedCallers is empty
  console.error("No allowed callers specified.");
}

// If the MicrosoftAppTenantId is specified in the environment config, add the tenant as a valid JWT token issuer for Bot to Skill conversation.
// The token issuer for MSI and single tenant scenarios will be the tenant where the bot is registered.
let validTokenIssuers = [];
const  MicrosoftAppTenantId  = '9b2bf913-dd6f-4c1f-bad4-46cd6522f9e5';//process.env.TEAMS_APP_TENANT_ID;
console.log("MSTenantid " + MicrosoftAppTenantId);

if (MicrosoftAppTenantId) {
    // For SingleTenant/MSI auth, the JWT tokens will be issued from the bot's home tenant.
    // Therefore, these issuers need to be added to the list of valid token issuers for authenticating activity requests.
    validTokenIssuers = [
        `${ AuthenticationConstants.ValidTokenIssuerUrlTemplateV1 }${ MicrosoftAppTenantId }/`,
        `${ AuthenticationConstants.ValidTokenIssuerUrlTemplateV2 }${ MicrosoftAppTenantId }/v2.0/`,
        `${ AuthenticationConstants.ValidGovernmentTokenIssuerUrlTemplateV1 }${ MicrosoftAppTenantId }/`,
        `${ AuthenticationConstants.ValidGovernmentTokenIssuerUrlTemplateV2 }${ MicrosoftAppTenantId }/v2.0/`
    ];
}

// Define our authentication configuration.
const authConfig = new AuthenticationConfiguration([], claimsValidators, validTokenIssuers);
console.log("authconfig " + authConfig);
/*const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
  new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: config.botId,
    MicrosoftAppPassword: process.env.BOT_PASSWORD,
    MicrosoftAppType: "MultiTenant",
  }),
  authConfig
); */

const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: config.botId,
  MicrosoftAppPassword: config.botPassword,
  MicrosoftAppType: "MultiTenant",
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  process.env,
  credentialsFactory,
  authConfig
);

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about how bots work.
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context, error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity("The bot encountered an error or bug.");
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;





// Listen for incoming server requests.
server.post("/api/messages", async (req, res) => {
  // Route received a request to adapter for processing
  await adapter.process(req, res as any, async (context) => {
    // Dispatch to application for routing
    await app.run(context);
  });
});
