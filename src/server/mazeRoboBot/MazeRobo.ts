import { BotDeclaration } from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, TeamsActivityHandler, MessageFactory, } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";
import * as Util from "util";
import ColourCard from "./dialogs/ColourDialog";
import * as request from "request-promise-native";
const TextEncoder = Util.TextEncoder;
// Initialize debug logging module
const log = debug("msteams");
const fetch = require('node-fetch');
/**
 * Implementation for Maze Robo
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_ID,
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_PASSWORD)

export class MazeRobo extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super();

        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));

        // Set up the Activity processing

        this.onMessage(async (context: TurnContext): Promise<void> => {
            // TODO: add your own bot logic in here
            switch (context.activity.type) {
                case ActivityTypes.Message:
                    if (context.activity.value) {
                        switch (context.activity.value.cardAction) {
                          case "turnOn":
                            await this.turnOnCardActivity(context);
                            break;
                          case "turnOff":
                            await this.turnOffCardActivity(context);
                            break;
                          case "changeColourPurple":
                            await this.changeColourPurpleCardActivity(context);
                            break;
                          case "changeColourTeal":
                            await this.changeColourTealCardActivity(context);
                            break;
                          case "changeColourOrange":
                            await this.changeColourOrangeCardActivity(context);
                            break;
                        }
                    } else {
                    let text = TurnContext.removeRecipientMention(context.activity);
                    text = text.toLowerCase();
                    if (text.startsWith("mentionme")){
                        await this.handleMessageMentionMeOneOnOne(context);
                        return;
                    } else if (text.startsWith("hello")) {
                        await context.sendActivity("Oh, hello to you as well!");
                        return;
                    } else if (text.startsWith("help")) {
                        const dc = await this.dialogs.createContext(context);
                        await dc.beginDialog("help");
                    } else if (text.startsWith("change colour")){
                        const colourCard = CardFactory.adaptiveCard(ColourCard);
                        await context.sendActivity({ attachments: [colourCard] });
                    } else if (text.startsWith("turn on")){
                        await this.handleMessageTurnOn(context);
                    } else if (text.startsWith("turn off")){
                        await this.handleMessageTurnOff(context);
                    } else {
                        await context.sendActivity(`I\'m terribly sorry, but my master hasn\'t trained me to do anything yet...`);
                    }
                    break;
                    }
                    default:
                    break;
            }
            // Save state changes
            return this.conversationState.saveChanges(context);
        });

        this.onConversationUpdate(async (context: TurnContext): Promise<void> => {
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                        await context.sendActivity({ attachments: [welcomeCard] });
                    }
                }
            }
        });

        this.onMessageReaction(async (context: TurnContext): Promise<void> => {
            const added = context.activity.reactionsAdded;
            if (added && added[0]) {
                await context.sendActivity({
                    textFormat: "xml",
                    text: `That was an interesting reaction (<b>${added[0].type}</b>)`
                });
            }
        });;
   }

   private async handleMessageMentionMeOneOnOne(context: TurnContext): Promise<void> {
    const mention = {
      mentioned: context.activity.from,
      text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
      type: "mention"
    };
  
    const replyActivity = MessageFactory.text(`Hi ${mention.text} from a 1:1 chat.`);
    replyActivity.entities = [mention];
    await context.sendActivity(replyActivity);
  }

  private async deleteCardActivity(context): Promise<void> {
    await context.deleteActivity(context.activity.replyToId);
  }

  private async handleMessageTurnOn(context): Promise<void> {
    const response = await fetch(process.env.URL, {
      method: 'PUT',
      body: '{"on": true, "hue": 2002, "sat":254}',
      headers: {'Content-Type': 'application/json; charset=UTF-8'} 
    });

    const replyActivity = MessageFactory.text(` Sure ${context.activity.from.name}, I have turned the light on.`);
    await context.sendActivity(replyActivity);
  }

  private async handleMessageTurnOff(context): Promise<void> {
    const response = await fetch(process.env.URL, {
      method: 'PUT',
      body: '{"on": false}',
      headers: {'Content-Type': 'application/json; charset=UTF-8'} 
    });

    const replyActivity = MessageFactory.text(`Sure ${context.activity.from.name}, I have turned the light off.`);
    await context.sendActivity(replyActivity);
  }

  private async changeColourPurpleCardActivity(context): Promise<void> {
    await context.deleteActivity(context.activity.replyToId);

    const response = await fetch(process.env.URL, {
      method: 'PUT',
      body: '{"on": true, "hue": 51245, "sat":254}',
      headers: {'Content-Type': 'application/json; charset=UTF-8'} 
    });

    const replyActivity = MessageFactory.text(`${context.activity.from.name} changed the colour to purple.`);
    await context.sendActivity(replyActivity);
  }

  private async changeColourTealCardActivity(context): Promise<void> {
    await context.deleteActivity(context.activity.replyToId);

    const response = await fetch(process.env.URL, {
      method: 'PUT',
      body: '{"on": true, "hue": 38775, "sat":254}',
      headers: {'Content-Type': 'application/json; charset=UTF-8'} 
    });

    const replyActivity = MessageFactory.text(`${context.activity.from.name} changed the colour to teal.`);
    await context.sendActivity(replyActivity);
  }

  private async changeColourOrangeCardActivity(context): Promise<void> {
    await context.deleteActivity(context.activity.replyToId);

    const response = await fetch(process.env.URL, {
      method: 'PUT',
      body: '{"on": true, "hue": 2002, "sat":254}',
      headers: {'Content-Type': 'application/json; charset=UTF-8'} 
    });

    const replyActivity = MessageFactory.text(`${context.activity.from.name} changed the colour to orange.`);
    await context.sendActivity(replyActivity);
  }

  private async turnOnCardActivity(context): Promise<void> {
    await context.deleteActivity(context.activity.replyToId);

    const response = await fetch(process.env.URL, {
      method: 'PUT',
      body: '{"on": true, "hue": 2002, "sat":254}',
      headers: {'Content-Type': 'application/json; charset=UTF-8'} 
    });

    var purple = '{"on": true, "hue": 51245, "sat":254}';
    var teal = '{"on": true, "hue": 38775, "sat":254}';
    var orange = '{"on": true, "hue": 2002, "sat":254}';
    const replyActivity = MessageFactory.text(`${context.activity.from.name} turned the light on.`);
    await context.sendActivity(replyActivity);
  }

  private async turnOffCardActivity(context): Promise<void> {
    await context.deleteActivity(context.activity.replyToId);

    const response = await fetch(process.env.URL, {
      method: 'PUT',
      body: '{"on": false}',
      headers: {'Content-Type': 'application/json; charset=UTF-8'} 
    });
    
    const replyActivity = MessageFactory.text(`${context.activity.from.name} turned the light off.`);
    await context.sendActivity(replyActivity);
  }

}