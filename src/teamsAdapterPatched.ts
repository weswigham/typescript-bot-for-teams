import { TeamsAdapter, TeamsContext, TeamsConnectorClient, Teams, TeamsChannelData } from "botbuilder-teams";
import { TurnContext, Activity, ConversationReference, ActivityTypes } from "botbuilder";
import * as msRest from "@azure/ms-rest-js";

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

const serializer = new msRest.Serializer({ TeamsCreateReplyChainRequest, CreateReplyChainCreatedResponse });
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
        return createReplyChain(teamsCtx, {
            activity: o,
            channelData: { channel: { id: ref.channelId } }
        });
    }
}