import { TeamsAdapter, TeamsContext, TeamsConnectorClient, Teams, TeamsChannelData } from "botbuilder-teams";
import { TurnContext, Activity, ConversationReference, ActivityTypes } from "botbuilder";
import * as msRest from "@azure/ms-rest-js";
import * as Mappers from "botbuilder-teams/lib/schema/models/mappers";
import { MicrosoftAppCredentials } from "botframework-connector";

async function createReplyChain(ctx: TeamsContext, args: { activity: Partial<Activity>, channelData: TeamsChannelData }) {
    return (await ctx.teamsConnectorClient.teams["client"].sendOperationRequest(
        {
            teamsCreateReplyChainRequest: args
        },
        createReplyChainOperationSpec
    ));
}

export const TeamsCreateReplyChainRequest: msRest.CompositeMapper = {
    serializedName: "teamsCreateReplyChainRequest",
    type: {
        name: "Composite",
        className: "TeamsCreateReplyChainRequest",
        modelProperties: {
            activity: {
                serializedName: "activity",
                type: {
                    name: "Composite",
                    className: "Activity"
                }
            },
            channelData: {
                serializedName: "channelData",
                type: {
                    name: "Composite",
                    className: "TeamsChannelData"
                }
            }
        }
    }
};

export const CreateReplyChainCreatedResponse: msRest.CompositeMapper = {
    serializedName: "CreateReplyChainCreatedResponse",
    type: {
        name: "Composite",
        className: "CreateReplyChainCreatedResponse",
        modelProperties: {
            id: {
                serializedName: "id",
                type: {
                    name: "String"
                }
            },
            activityId: {
                serializedName: "activityId",
                type: {
                    name: "String"
                }
            }
        }
    }
};

// THIS TYPE ALLOWS FOR REPRESENTING SO MANY INVALID STATES
// AND THIS ONLY EVEN ENCODES SOME OF THE FIELDS FROM
// https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-api-reference
const Activity: msRest.CompositeMapper = {
    serializedName: "Activity",
    type: {
        name: "Composite",
        className: "Activity",
        modelProperties: {
            action: {
                serializedName: "action",
                type: {
                    name: "String"
                }
            },
            text: {
                serializedName: "text",
                type: {
                    name: "String"
                }
            },
            type: {
                serializedName: "type",
                type: {
                    name: "String"
                }
            },
        }
    }
};

const serializer = new msRest.Serializer({ Activity, TeamsCreateReplyChainRequest, CreateReplyChainCreatedResponse, ...Mappers });
const createReplyChainOperationSpec: msRest.OperationSpec = {
    httpMethod: "POST",
    path: "v3/conversations",
    requestBody: {
        parameterPath: "teamsCreateReplyChainRequest",
        mapper: {
            ...TeamsCreateReplyChainRequest,
            required: true
        }
    },
    responses: {
        201: {
            bodyMapper: CreateReplyChainCreatedResponse
        },
        default: {}
    },
    serializer
};

export class PatchedTeamsAdapter extends TeamsAdapter {
    createReplyChainFromConversationReference(context: TurnContext, ref: Partial<ConversationReference>, activity: Partial<Activity>) {
        const o: Partial<Activity> = TurnContext.applyConversationReference({ ...activity }, ref);
        if (!o.type) { o.type = ActivityTypes.Message; }

        const teamsCtx = TeamsContext.from(context);
        const hardcodedServiceUrl = "https://smba.trafficmanager.net/amer/";
        console.log(`Trusting service URL${!activity.serviceUrl ? ref.serviceUrl ? " (from ref)" : " (hardcoded)" : ""} ${activity.serviceUrl || ref.serviceUrl || hardcodedServiceUrl}`);
        MicrosoftAppCredentials.trustServiceUrl(activity.serviceUrl || ref.serviceUrl || hardcodedServiceUrl);
        return createReplyChain(teamsCtx, {
            activity: o,
            channelData: { channel: { id: ref.conversation.id.split(";")[0] } }
        });
    }
}