// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { TurnContext, BotState, ConversationReference, Storage } from 'botbuilder';
import { CronJob, CronTime } from "cron";
import moment = require("moment-timezone");
import { PatchedTeamsAdapter } from './teamsAdapterPatched';

function validate(cronExp: string) {
    try {
        new CronTime(cronExp);
        return true;
    }
    catch {}
    return false;
}

const scheduleMessageRegex = /schedule standup here on (.*)/;

interface BotStorageSchema {
    cronExpr: { value: string, eTag: string; };
    ref: { value: Partial<ConversationReference>, eTag: string; };
}

export class TypescriptStandupBot {
    private task: CronJob | undefined;
    constructor(
        private conversationState: BotState,
        private storage: Storage,
        private adapter: PatchedTeamsAdapter
    ) {
        // This is potentially racy if a message arives and is handled before this is done
        // If that turns out to be problematic, we can try to queue turns while initialization
        // is ongoing - but that's quite a bit for an edge case that might not be a problem.
        this.setupAsync();
    }

    async setupAsync() {
        const data: Partial<BotStorageSchema> = await this.storage.read(["ref", "cronExpr"]);
        if (data.cronExpr && data.ref) {
            await this.setupStandupCron(data.cronExpr.value, data.ref.value, /*save*/ false);
        }
    }

    async postStandupThread(ref: Partial<ConversationReference>) {
        this.adapter.continueConversation(ref, async context => {
            const date = moment().tz("America/Los_Angeles").format("YYYY-MM-DD");
            await this.adapter.createReplyChainFromConversationReference(context, ref, {
                text: `**Standup ${date}**`
            });
        });
    }

    async setupStandupCron(cronExpr: string, ref: Partial<ConversationReference>, save: boolean) {
        if (this.task) {
            this.task.stop();
        }
        this.task = new CronJob(cronExpr, () => {
            this.postStandupThread(ref);
        }, null, true, "America/Los_Angeles", this);
        if (save) {
            const toStore: BotStorageSchema = {
                cronExpr: { value: cronExpr, eTag: "*" },
                ref: { value: ref, eTag: "*" }
            };
            await this.storage.write(toStore);
        }
    }

    async handleSceduleStandupMessage(turnContext: TurnContext, cronExpr: string) {
        if (!validate(cronExpr)) {
            await turnContext.sendActivity(`'${cronExpr}' is not a valid cron expression. Try something like '30 10 * * mon,tue,wed,thu,fri', which would schedule a standup at 10:30 each weekday.`);
            return;
        }
        const ref = TurnContext.getConversationReference(turnContext.activity);
        if (this.task) {
            await turnContext.sendActivity("Existing standup schedule canceled, resceduling...");
        }
        await this.setupStandupCron(cronExpr, ref, /*save*/ true);
        await turnContext.sendActivity(`Standups scheduled for '${cronExpr}', next up on ${this.task!.nextDate().toString()}`);
    }

    async onTurn(turnContext: TurnContext) {
        if (turnContext.activity.type === "message") {
            const text = turnContext.activity.text;
            const match = scheduleMessageRegex.exec(text);
            if (match) {
                await this.handleSceduleStandupMessage(turnContext, match[1]);
            }
        }
        await this.conversationState.saveChanges(turnContext);
    }
}
