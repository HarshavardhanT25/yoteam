import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult, TaskModuleRequest, TaskModuleContinueResponse } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";

// Initialize debug logging module
const log = debug("msteams");

// export const USERID = `${process.env.REACT_APP_INSTANCE_USERID}`;
// export const PASSWORD = `${process.env.REACT_APP_INSTANCE_PASSWORD}`;
// export const INSTANCE = `${process.env.REACT_APP_INSTANCE}`;

@PreventIframe("/yoteamMessageExtension/config.html")
@PreventIframe("/yoteamMessageExtension/action.html")
export default class YoteamMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onFetchTask(context: TurnContext, value: MessagingExtensionQuery): Promise<MessagingExtensionResult | TaskModuleContinueResponse> {
        return Promise.resolve<TaskModuleContinueResponse>({
            type: "continue",
            value: {
                url: `https://${process.env.PUBLIC_HOSTNAME}/yoteamMessageExtension/action.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`,
                height: "large"
            }
        });

    }

    // handle action response in here
    // See documentation for `MessagingExtensionResult` for details
    public async onSubmitAction(context: TurnContext, value: TaskModuleRequest): Promise<MessagingExtensionResult> {
        console.log(value);
        return Promise.resolve({
            type: "message",
            text: ""
        });
    }

    // this is used when canUpdateConfiguration is set to true
    public async onQuerySettingsUrl(context: TurnContext): Promise<{ title: string, value: string }> {
        return Promise.resolve({
            title: "yoteam Message Extension Configuration",
            value: `https://${process.env.PUBLIC_HOSTNAME}/yoteamMessageExtension/config.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`
        });
    }

    public async onSettings(context: TurnContext): Promise<void> {
        // take care of the setting returned from the dialog, with the value stored in state
        const setting = context.activity.value.state;
        log(`New setting: ${setting}`);
        return Promise.resolve();
    }

}
