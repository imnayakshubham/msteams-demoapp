import { default as axios } from "axios";
import * as querystring from "querystring";
import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionAction,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  MessagingExtensionActionResponse,
  AppBasedLinkQuery,
  TeamsInfo,
  MessageFactory,
  InvokeResponse,
  BotHandler,
  BotFrameworkAdapter,
  ActivityTypes,
} from "botbuilder";
import { ConnectorClient, MicrosoftAppCredentials } from 'botframework-connector';
import config from "./config";

const connectorClient = new ConnectorClient(new MicrosoftAppCredentials(config.botId, config.botPassword));

const PORT = process.env.port || process.env.PORT || 3978
const endPoint = `https://localhost:${PORT}/api/create-conversion`

async function sendMessageToPersonalChat(userData) {
  try {
    // Create a new conversation with the user in their personal chat
    const resource = await connectorClient.conversations.createConversation({
      bot: { id: config.botId, name: "sendmessage" },
      members: [{ id: userData.id, name: userData.name } // Provide the user ID as the recipient
      ],
      isGroup: false,
      activity: null,
      channelData: null
    });

    // Create an Activity object for the message
    const activity = {
      type: 'message',
      text: 'Hello, this is a message from the bot!',
      conversation: { id: resource.id, isGroup: false, conversationType: null, name: null }
    };

    // Send the message to the user's personal chat
    await connectorClient.conversations.sendToConversation(resource.id, activity);

    console.log('Message sent successfully to personal chat!');
  } catch (error) {
    console.error('Error sending message to personal chat:', error);
  }
}


const getAccessToken = async (tenantId) => {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`

  const postData = {
    grant_type: 'authorization_code',
    client_id: "bc9aa6cc-e8eb-43fa-8c69-5e83d30a4060",
    scope: "https://graph.microsoft.com/.default",
    code: null,
    redirect_uri: null
  };

  const data = querystring.stringify({
    'grant_type': 'client_credentials',
    'client_id': 'bc9aa6cc-e8eb-43fa-8c69-5e83d30a4060',
    'client_secret': 'VX98Q~KbLKhoAxfJ_L-Psmlt2Okl4wGB6HIm2cqs',
    'scope': 'https://graph.microsoft.com/.default'
  });

  const config = {
    method: 'post',
    url: url,
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      // 'Cookie': 'fpc=AsUV6XmErQVDkGj0D05NGYwlmCaFAQAAAKX3A9wOAAAA; stsservicecookie=estsfd; x-ms-gateway-slice=estsfd'
    },
    data: data
  };

  try {
    const response = await axios.post(config.url, config.data, { headers: config.headers });
    return response.data.access_token

  } catch (error) {
    console.error('Error retrieving access token:', error.response.data);
  }
}

const getUserEmailAddress = async (name, access_token) => {
  // Make a request to the Microsoft Graph API to retrieve the user's email address
  try {
    const response = await axios.post(`https://graph.microsoft.com/v1.0/users?$filter=displayNameeq'${name}'&$select=mail`, {
      headers: {
        Authorization: "Bearer " + access_token, // Replace with the appropriate access token for the Microsoft Graph API
      },
    });

  } catch (error) {
    console.log({ error });

  }

  // const data = await response
  // const user = data.value[0];

  // if (user && user.mail) {
  //   console.log(user);
  //   return user.mail;
  // } else {
  //   throw new Error("Failed to retrieve user email address");
  // }
};


export class TeamsBot extends TeamsActivityHandler {
  // Action.

  constructor() {
    super();

    // Register event handlers
    this.onMessage(async (context, next) => {
      // Handle incoming message activities
      // await this.handleMessageActivity(context);
      await next();
    });

    this.onMessageReaction(async (context, next) => {
      console.log({ contexts: context });
      // if (context.["activity"].reactions) {
      //   for (const reaction of context.activity.reactions) {
      //     if (reaction.type === 'messagingExtension/submitAction') {
      //       const action = reaction.action; // Extract the action data from the adaptive card button

      //       if (action === 'submit') {
      //         await context.sendActivity('Submit button clicked!'); // Perform specific action for the submit button
      //       } else if (action === 'openUrl') {
      //         await context.sendActivity('Open URL button clicked!'); // Perform specific action for the openUrl button
      //       }
      //     }
      //   }
      // }

      await next();
    });
  }



  public async handleTeamsMessagingExtensionFetchTask(context: TurnContext, action: any): Promise<MessagingExtensionActionResponse> {
    const adaptiveCard = CardFactory.adaptiveCard({
      actions: [{
        data: { submitLocation: 'messagingExtensionFetchTask' },
        title: 'Submit',
        type: 'Action.Submit'
      }],
      body: [
        { type: 'TextBlock', text: 'How can we improve the Work/Life Balance?', id: "question" },
        { id: 'question', placeholder: 'Your Suggestion', type: 'Input.Text', require: true },
      ],
      type: 'AdaptiveCard',
      version: '1.0'
    });

    return {
      task: {
        type: 'continue',
        value: {
          card: adaptiveCard,
          height: 200,
          title: 'Suggestion',
          url: null,
          width: 600
        }
      }
    };
  }

  public async handleTeamsMessagingExtensionSubmitAction(
    context: TurnContext,
    action: MessagingExtensionAction
  ): Promise<MessagingExtensionActionResponse> {
    switch (action.commandId) {
      case "openModal":
        return createCardCommand(context, action);
      case "shareMessage":
        return shareMessageCommand(context, action);
      default:
        throw new Error("NotImplemented");
    }
  }

  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    const searchQuery = query.parameters[0].value;
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
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }

  public async handleTeamsMessagingExtensionSelectItem(
    context: TurnContext,
    obj: any
  ): Promise<MessagingExtensionResponse> {
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [CardFactory.heroCard(obj.name, obj.description)],
      },
    };
  }

  // Link Unfurling.
  public async handleTeamsAppBasedLinkQuery(
    context: TurnContext,
    query: AppBasedLinkQuery
  ): Promise<MessagingExtensionResponse> {
    const attachment = CardFactory.thumbnailCard("Image Preview Card", query.url, [query.url]);

    // By default the link unfurling result is cached in Teams for 30 minutes.
    // The code has set a cache policy and removed the cache for the app. Learn more here: https://learn.microsoft.com/microsoftteams/platform/messaging-extensions/how-to/link-unfurling?tabs=dotnet%2Cadvantages#remove-link-unfurling-cache
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [attachment],
        suggestedActions: {
          actions: [
            {
              title: "default",
              type: "setCachePolicy",
              value: '{"type":"no-cache"}',
            },
          ],
        },
      },
    };
  }
}

function getTeamId(context: TurnContext) {
  try {
    const teamId = context?.["_activity"].channelData.tenant.id;
    return teamId;
  } catch (error) {
    console.error(`Error retrieving Team ID: ${error}`);
    return null;
  }
}


const sendAdaptiveCardToChannel = async (accessToken, teamId, channelId) => {
  try {
    const payload = {
      "type": "message",
      "attachments": [
        {
          "contentType": "application/vnd.microsoft.card.adaptive",
          "content": {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.4",
            "body": [
              {
                "type": "TextBlock",
                "text": "Hello, Adaptive Card!",
                "size": "Large"
              }
            ]
          }
        }
      ]
    };


    // const messagePayload = {
    //   body: {
    //     contentType: 'html',
    //     content: `<attachment id="adaptiveCardAttachment"></attachment>`,
    //   },
    //   attachments: [
    //     {
    //       id: 'adaptiveCardAttachment',
    //       contentType: 'application/vnd.microsoft.card.adaptive',
    //       content: adaptiveCard.content,
    //     },
    //   ],
    // };

    const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages`;

    const response = await axios.post(url, payload, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    });

    console.log('Adaptive Card sent successfully!');
    console.log('Message ID:', response.data.id);
  } catch (error) {
    console.error('Error sending Adaptive Card:', error.response?.data || error.message);
  }
};


async function sendUserData(context: TurnContext, userData: any): Promise<any> {
  const teamId = "cb1cf1c0-ebb5-41df-bd66-136e7566c4dc" || getTeamId(context)
  const channelId = "19:2859d8fb4ddf4ce78c592a5c235df803@thread.tacv2"
  const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages`;
  const accessToken = await getAccessToken(context?.["_activity"]?.channelData.tenant.id)

  const generalChannelId = '19:lFBaEpStgzTEEbPqIXx4OMhWlYsJVpQl--GvH07DG1A1@thread.tacv2'

  const adaptiveCard = {
    actions: [{
      data: { submitLocation: 'messagingExtensionFetchTask' },
      title: 'Acknowledge',
      type: 'Action.Submit'
    }],
    body: [
      { type: 'TextBlock', text: 'How can we improve the Work/Life Balance?', id: "question" },
      { type: 'TextBlock', text: userData.question, id: "question" },
    ],
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    type: 'AdaptiveCard',
    version: '1.0'
  };

  // const adaptiveCardAttachment = {
  //   contentType: 'application/vnd.microsoft.card.adaptive',
  //   content: adaptiveCard
  // };
  // const message = MessageFactory.attachment(adaptiveCardAttachment);
  // message.channelId = channelId;
  // message.type = ActivityTypes.Message;


  // await context.sendActivity(message);

  // const payload = {
  //   "body": {
  //     "contentType": "html",
  //     "content": "<attachment id=\"4465B062-EE1C-4E0F-B944-3B7AF61EAF40\"></attachment>"
  //   },
  //   "attachments": [
  //     {
  //       "id": "4465B062-EE1C-4E0F-B944-3B7AF61EAF40",
  //       "contentType": "application/vnd.microsoft.card.adaptive",
  //       "content": `{\"type\": \"AdaptiveCard\",\"$schema\": \http://adaptivecards.io/schemas/adaptive-card.json\,\"version\": \"1.3\",\"body\": [  {\"type\": \"TextBlock\",\"size\": \"Large\",\"weight\": \"Bolder\",\"text\": \"My News Item\",\"wrap\": true  }],\"actions\": [  {\"type\": \"Action.Execute\",\"title\": \"View\",\"url\": \https://bing.com\  }]  }`
  //     }
  //   ]
  // }


  // const payload = {
  //   contentType: "application/vnd.microsoft.card.adaptive",
  //   content: adaptiveCard,
  // };

  // console.log({
  //   "user": {
  //     ...context?.["_activity"]?.from,
  //     "userIdentityType": "aadUser"
  //   }
  // });

  // const data = {
  //   // "createdDateTime": "2019-02-04T19:58:15.511Z",
  //   "from": {
  //     "user": {
  //       id: context?.["_activity"]?.from.id,
  //       name: context?.["_activity"]?.from.name,
  //       "userIdentityType": "aadUser"
  //     }
  //   },
  //   body: {
  //     // contentType: 'html',
  //     content: '<attachment id="4465B062-EE1C-4E0F-B944-3B7AF61EAF40"></attachment>'
  //   },
  //   attachments: [
  //     {
  //       id: '4465B062-EE1C-4E0F-B944-3B7AF61EAF40',
  //       contentType: 'application/vnd.microsoft.card.adaptive',
  //       content: JSON.stringify({
  //         type: 'AdaptiveCard',
  //         $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
  //         version: '1.3',
  //         body: [
  //           {
  //             type: 'TextBlock',
  //             size: 'Large',
  //             weight: 'Bolder',
  //             text: 'How can we improve the Work/Life Balance?',
  //             wrap: true
  //           },
  //           {
  //             type: 'TextBlock',
  //             size: 'Large',
  //             weight: 'Bolder',
  //             text: '${userData.question}',
  //             wrap: true
  //           }
  //         ],
  //         actions: [
  //           {
  //             type: 'Action.Execute',
  //             title: 'Acknowledge',
  //             url: 'https://bing.com'
  //           }
  //         ]
  //       })
  //     }
  //   ]
  // };

  // const config = {
  //   method: 'post',
  //   url: url,
  //   headers: {
  //     Authorization: `Bearer ${accessToken}`,
  //   },
  //   data: data
  // };

  // console.log(config.url, config.data, { headers: config.headers });
  // const response = await sendAdaptiveCardToChannel(accessToken, teamId, channelId);
  // console.log({ response });

  const incomingWebhook = `https://7xlg42.webhook.office.com/webhookb2/cb1cf1c0-ebb5-41df-bd66-136e7566c4dc@cd325ff4-ffae-44da-910e-dc88dd28a4df/IncomingWebhook/7ce59b8b4315426bbcf9a2c7cb98064e/5185948e-247d-4cd8-99e1-6bdce640776b`

  const cardPayload = {
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    "themeColor": "0072C6",
    "summary": "Demo Adaptive Card",
    "sections": [
      {
        "activityTitle": "Question: How can we improve the Work/Life Balance?",
        "activitySubtitle": "Answer: " + userData.question,
        "facts": [

        ],
        "markdown": true
      },
    ],
    "potentialAction": [
      {
        "@type": "ActionCard",
        "name": "Acknowledge",
        "actions": [
          {
            "@type": "HttpPOST",
            "name": "Acknowledge",
            "target": endPoint,
            "headers": [
              {
                "name": "Content-Type",
                "value": "application/json"
              }
            ],
            "body": "{\"email\":\"demo1@demo.com\",\"password\":\"b05e15c2388dafff9547831ae865b92e8993a47c\",\"user_name\":\"test@qe\"}"
          }
        ]
      }
    ]

  };

  try {
    const response = await axios.post(incomingWebhook, cardPayload);
    // console.log({ response });
    // if (response) {
    //   await sendMessageToPersonalChat(userData.posted_by)
    // }

  } catch (error) {
    console.error('Error sending message:', error.response.data);
  }
}


async function createCardCommand(
  context: TurnContext,
  action: MessagingExtensionAction
) {
  // The user has chosen to create a card by choosing the 'Create Card' context menu command.
  const data = action.data;

  const response = await sendUserData(context, { ...data, posted_by: context["_activity"].from })
  return null
}

async function shareMessageCommand(
  context: TurnContext,
  action: MessagingExtensionAction
): Promise<MessagingExtensionResponse> {
  // The user has chosen to share a message by choosing the 'Share Message' context menu command.
  let userName = "unknown";
  if (
    action.messagePayload &&
    action.messagePayload.from &&
    action.messagePayload.from.user &&
    action.messagePayload.from.user.displayName
  ) {
    userName = action.messagePayload.from.user.displayName;
  }

  // This Message Extension example allows the user to check a box to include an image with the
  // shared message.  This demonstrates sending custom parameters along with the message payload.
  let images = [];
  const includeImage = action.data.includeImage;
  if (includeImage === "true") {
    images = [
      "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU",
    ];
  }
  const heroCard = CardFactory.heroCard(
    `${userName} originally sent this message:`,
    action.messagePayload.body.content,
    images
  );

  if (
    action.messagePayload &&
    action.messagePayload.attachments &&
    action.messagePayload.attachments.length > 0
  ) {
    // This sample does not add the MessagePayload Attachments.  This is left as an
    // exercise for the user.
    heroCard.content.subtitle = `(${action.messagePayload.attachments.length} Attachments not included)`;
  }

  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}
