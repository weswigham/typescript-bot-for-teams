import { TeamsAdapter, TeamsContext, TeamsConnectorClient, Teams, TeamsChannelData } from "botbuilder-teams";
import { TurnContext, Activity, ConversationReference, ActivityTypes } from "botbuilder";
import * as msRest from "@azure/ms-rest-js";
import * as Mappers from "botbuilder-teams/src/schema/models/mappers";

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

const serializer = new msRest.Serializer({ ...Mappers, TeamsCreateReplyChainRequest, CreateReplyChainCreatedResponse });
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
    createReplyChainPatched(turnContext: TurnContext, activity: Partial<Activity>, inGeneralChannel?: boolean) {
        let sentNonTraceActivity: boolean = false;
        const teamsCtx = TeamsContext.from(turnContext);
        const ref: Partial<ConversationReference> = TurnContext.getConversationReference(turnContext.activity);
        const o: Partial<Activity> = TurnContext.applyConversationReference({ ...activity }, ref);
        try {
            o.conversation.id = inGeneralChannel
                ? teamsCtx.getGeneralChannel().id
                : teamsCtx.channel.id;
        } catch (e) {
            // do nothing for fields fetching error
        }
        if (!o.type) { o.type = ActivityTypes.Message; }
        if (o.type !== ActivityTypes.Trace) { sentNonTraceActivity = true; }

        return createReplyChain(teamsCtx, {
            activity: o,
            channelData: teamsCtx.getTeamsChannelData()
        }).then(resp => {
            // Set responded flag
            if (sentNonTraceActivity) { turnContext.responded = true; }
            return resp;
        });
    }
}