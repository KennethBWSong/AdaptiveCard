import { ActivityHandler, ActivityTypes, CardFactory, InvokeResponse, TeamsActivityHandler, TurnContext } from "botbuilder";
import config from "./config";
import { OnBehalfOfUserCredential } from "@microsoft/teamsfx";

export class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context, next) => {
      await context.sendActivity("Start login");
      console.log("Running with Message Activity.");
      const card = this.createViewProfileCard();
      const response = CardFactory.adaptiveCard(card);
      await context.sendActivity({ attachments: [response,] });
      await next();
    });
    this.onMembersAdded(async (context, next) => {
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });
  }

  async onInvokeActivity(context: TurnContext) {
    if (
      context.activity.name != "adaptiveCard/action" || 
      !context.activity.value ||
      !context.activity.value["action"]
    ) {
      return null;
    }

    const value = context.activity.value;
    const authentication = value["authentication"] ?? null;
    const token = authentication?.token;

    if (!token) {
      const card = this.initialSso(context);
      return ActivityHandler.createInvokeResponse(card);
    } else {
      await context.sendActivity(`Token: ${token}`);
      await this.tokenExchange(token, context);
    }
  }

  initialSso(context: TurnContext) {
    const signInLink = config.loginStartPage + `?clientId=${config.clientId}&tenantId=${config.tenantId}&scope=User.Read`;
    console.log(signInLink);

    const response = {
      statusCode: 401,
      type: "application/vnd.microsoft.activity.loginRequest",
      value: {
        text: "SignIn Text",
        connectionName: "Test",
        buttons: [
          {
            title: "Sign-In",
            text: "Sign-In",
            type: "signin",
            value: signInLink,
          }
        ],
        tokenExchangeResource: {
          id: config.clientId,
          uri: config.appIdUri,
        }
      }
    };
    return response;
  }

  consent(id: string) {
    return {
      type: "invokeResponse",
      value: {
        status: 412,
        body: {
          id: id,
          failureDetail: "The bot is unable to exchange token. Ask for user consent."
        }
      }
    }
  }

  createViewProfileCard() {
    return {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "version": "1.5",
      "type": "AdaptiveCard",
      "body": [
        {
          "type": "TextBlock",
          "text": "",
          "size": "Medium",
          "weight": "Bolder"
        },
        {
          "type": "ActionSet",
          "fallback": "drop",
          "actions": [
            {
              "type": "Action.Execute",
              "title": "Sign in to view profile'",
              "verb": "signin"
            }
          ]
        }
      ]
    }
  }

  async tokenExchange(ssoToken: string, context: TurnContext) {
    const credential = new OnBehalfOfUserCredential(ssoToken, {
      tenantId: config.tenantId,
      clientId: config.clientId,
      initiateLoginEndpoint: config.loginStartPage,
      applicationIdUri: config.appIdUri,
      authorityHost: "https://login.microsoftonline.com",
      clientSecret: config.clientSecret
    });
    try {
      const graphToken = await credential.getToken("User.Read");
      await context.sendActivity(`Graph: ${graphToken.token}`);
    } catch (error) {
      await context.sendActivity({
        value: {
            body: {
                statusCode: 412,
                type: 'application/vnd.microsoft.error.preconditionFailed',
                value: {
                    code: '412',
                    message: 'Failed to exchange token'
                }
            },
            status: 200
        } as InvokeResponse,
        type: ActivityTypes.InvokeResponse
      });
      return undefined;
    }
  }
}
