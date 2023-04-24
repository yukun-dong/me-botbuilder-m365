// Import required packages
import * as restify from "restify";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { CardFactory, CloudAdapter, ConfigurationBotFrameworkAuthentication, ConfigurationServiceClientCredentialFactory, TurnContext } from "botbuilder";

// This bot's main dialog.
import config from "./config";
import { default as axios } from "axios";
import * as querystring from "querystring";

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: config.botId,
  MicrosoftAppPassword: config.botPassword,
  MicrosoftAppType: "MultiTenant",
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
  credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
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
  await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Set the onTurnError for the singleton BotFrameworkAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create the bot that will handle incoming messages.
// const bot = new TeamsBot();

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});



import {
  Application,
  DefaultConversationState,
  DefaultPromptManager,
  DefaultTempState,
  DefaultTurnState,
  DefaultUserState,
  OpenAIPlanner,
  RouteSelector
} from '@microsoft/botbuilder-m365';

const app = new Application({ adapter, botAppId: config.botId });

//link unfurling
const routeSelector: RouteSelector = async (context: TurnContext) => {
  if (context.activity.value.url && context.activity.name === 'composeExtension/queryLink') return Promise.resolve(true)
  else return Promise.resolve(false)
};
app.messageExtensions.queryLink(routeSelector, async (context: TurnContext, state: DefaultTurnState) => {
  const card = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.5",
    "body": [
      {
        "type": "TextBlock",
        "text": "link unfurling",
        "size": "large",
        "wrap": true,
        "style": "heading"
      },
    ]
  }

  const attachment = { ...CardFactory.adaptiveCard(card), preview: CardFactory.heroCard("test", "test") };
  return {
    type: "result",
    attachmentLayout: "list",
    attachments: [attachment],
  };
})

//zero install link unfurling
const zeroInstallRouteSelector: RouteSelector = async (context: TurnContext) => {
  if (context.activity.value.url && context.activity.name === 'composeExtension/anonymousQueryLink') return Promise.resolve(true)
  else return Promise.resolve(false)
};
app.messageExtensions.anonymousQueryLink(zeroInstallRouteSelector, async (context: TurnContext, state: DefaultTurnState) => {
  const card = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.5",
    "body": [
      {
        "type": "TextBlock",
        "text": "zero install link unfurling",
        "size": "large",
        "wrap": true,
        "style": "heading"
      },
    ]
  }

  const attachment = { ...CardFactory.adaptiveCard(card), preview: CardFactory.heroCard("test", "test") };
  return {
    type: "result",
    attachmentLayout: "list",
    attachments: [attachment],
  };
})


//search
app.messageExtensions.query("searchQuery", async (context: TurnContext, state: DefaultTurnState, query: any) => {
  const searchQuery = query.parameters.searchQuery;
  const response = await axios.get(
    `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
      text: searchQuery,
      size: 8,
    })}`
  );

  const attachments = [];
  response.data.objects.forEach((obj) => {
    const heroCard = CardFactory.heroCard(obj.package.name);
    const preview = CardFactory.heroCard(obj.package.name);
    preview.content.tap = {
      type: "invoke",
      value: { name: obj.package.name, description: obj.package.description },
    };
    const attachment = { ...heroCard, preview };
    attachments.push(attachment);
  });
  return {
    type: "result",
    attachmentLayout: "list",
    attachments: attachments,
  }
});

app.messageExtensions.selectItem( async (context: TurnContext, state: DefaultTurnState, item: any) => {
  const data = item;
  const heroCard = CardFactory.heroCard(data.name, data.description);
  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    type: "result",
    attachmentLayout: "list",
    attachments: [attachment],
  };
});


//action
app.messageExtensions.submitAction("createCard", async (context: TurnContext, state: DefaultTurnState, action: any) => {
  const data = action;
  const heroCard = CardFactory.heroCard(data.title, data.text);
  heroCard.content.subtitle = data.subTitle;
  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    type: "result",
    attachmentLayout: "list",
    attachments: [attachment],
  };
});

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await app.run(context);
  });
});
