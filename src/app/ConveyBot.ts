import { BotDeclaration } from "express-msteams-host";
import * as debug from "debug";
import { 
    CardFactory,
    TurnContext,
    MemoryStorage,
    ConversationState,
    TeamsActivityHandler,
    MessagingExtensionAction,
    MessagingExtensionActionResponse,
    MessageFactory,
    TaskModuleContinueResponse} from "botbuilder";

// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for Convey
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)

export class Convey extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super(); 
        this.conversationState = conversationState;

        // Set up the Activity processing
        this.onMessage(async (context: TurnContext, next): Promise<void> => {
            // TODO: add your own bot logic in here
            await context.sendActivity(`🦖Hi there. please use message extension`);
            await next();
        });        
    }

    protected async handleTeamsMessagingExtensionFetchTask(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
        log(`Event triggered - Fetch/Task`);

        return Promise.resolve({
            task: {
                width: "medium",
                height: "medium",
                type: "continue",
                value: {
                    title: "Input form",
                    url: `https://${process.env.HOSTNAME}/conveyMessageMessageExtension/action.html`
                }
            } as TaskModuleContinueResponse
        } as MessagingExtensionActionResponse);
    }

    protected async handleTeamsMessagingExtensionSubmitAction(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
        
        log(`Event triggered - SubmitAction`);
        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: "Hello User!"
                    },
                    {
                        type: "Image",
                        url: `https://randomuser.me/api/portraits/thumb/women/${Math.round(Math.random() * 100)}.jpg`
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.0"
            });

        await context.sendActivity(MessageFactory.attachment(card));
        // Send empty object
        return Promise.resolve({});
    }
}
