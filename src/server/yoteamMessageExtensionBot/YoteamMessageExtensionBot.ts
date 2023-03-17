import { BotDeclaration, MessageExtensionDeclaration } from "express-msteams-host";
import * as debug from "debug";
import { TeamsActivityHandler, StatePropertyAccessor, ActivityTypes, CardFactory, ConversationState, MemoryStorage, TurnContext } from "botbuilder";

import WelcomeCard from "./cards/welcomeCard";
import YoteamMessageExtension from "../yoteamMessageExtension/YoteamMessageExtension";
import { DialogSet, DialogState } from "botbuilder-dialogs";
// Initialize debug logging module
// const log = debug("msteams");

/**
 * Implementation for yoteam Message Extension Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_ID,
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_PASSWORD)

export class YoteamMessageExtensionBot extends TeamsActivityHandler {

    private readonly conversationState: ConversationState;
    /** Local property for YoteamMessageExtension */
    @MessageExtensionDeclaration("yoteamMessageExtension")
    private _yoteamMessageExtension: YoteamMessageExtension;

    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super();
        // Message extension YoteamMessageExtension
        this._yoteamMessageExtension = new YoteamMessageExtension();

        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        // Set up the Activity processing
        this.onMessage(async (context: TurnContext): Promise<void> => {
            // TODO: add your own bot logic in here
            switch (context.activity.type) {
                case ActivityTypes.Message:
                    {
                        let text = TurnContext.removeRecipientMention(context.activity);
                        text = text.toLowerCase();
                        if (text.startsWith("hello")) {
                            await context.sendActivity("Oh, hello to you as well!");
                            return;
                        } else if (text.startsWith("help")) {
                            await context.sendActivity("Please refer to [this link](https://docs.microsoft.com/en-us/microsoftteams/platform/bots/what-are-bots) to see how to develop bots for Teams");
                        } else {
                            await context.sendActivity("I'm terribly sorry, but my developer hasn't trained me to do anything yet...");
                        }
                    }
                    break;
                default:
                    break;
            }
            // Save state changes
            return this.conversationState.saveChanges(context);
        });
    }
}
