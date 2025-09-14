import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, IActionArguments, IFocusParameters } from '@microsoft/sp-adaptive-card-extension-base';
import {
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState
} from '../SpfxCardAdaptiveCardExtension';


export class InstanceQuickView extends BaseAdaptiveCardQuickView<
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState,
  //eslint-disable-next-line @typescript-eslint/no-explicit-any
  any
> {
  VIEW_ID = "UNKNOWN_VIEWID";
  errorMessage = "";
  errorCode = "";

  constructor(viewId: string) {
    super();
    this.VIEW_ID = viewId;
  }
  getErrorMessage(methodWhereError: string) {
    return `App ${this.state.selectedApp} registered QuickView ${this.VIEW_ID}, but the instance does not have a ${methodWhereError} method or it did not return anything.`;
  }
  //eslint-disable-next-line @typescript-eslint/no-explicit-any
  public get data(): any {
    const data = this.state.SPFxExtensionInstance?.getQuickViewData?.(this.VIEW_ID);
    if (data) return data;
    this.errorMessage = this.getErrorMessage("getQuickViewData");
    this.errorCode = "NO_GET_QUICKVIEW_DATA";
    return {};
  }

  public get focusParameters(): IFocusParameters | undefined {
    return this.state.SPFxExtensionInstance?.onQuickViewFocusParameter?.(this.VIEW_ID);
  }

  onAction(action: IActionArguments): void {
    this.state.SPFxExtensionInstance?.onAction?.(this.VIEW_ID, action);
  }

  public get template(): ISPFxAdaptiveCard {
    const template = this.state.SPFxExtensionInstance?.getQuickViewTemplate?.(this.VIEW_ID);
    if (template && !this.errorCode && !this.errorMessage) return template;

    const codeToUse = this.errorCode ? this.errorCode : "NO_GET_QUICKVIEW_TEMPLATE";
    const errorToUse = this.errorMessage ? this.errorMessage : this.getErrorMessage("getQuickViewTemplate");
    // Return a default template if none is provided by the SPFxExtensionInstance
    return {
      "schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.5",
      "body": [
        {
          "type": "TextBlock",
          "weight": "Bolder",
          "text": "Error!"
        },
        {
          "type": "ColumnSet",
          "columns": [
            {
              "type": "Column",
              "items": [
                {
                  "type": "TextBlock",
                  "weight": "Bolder",
                  "text": errorToUse,
                  "wrap": true
                },
                {
                  "type": "ActionSet",
                  "actions": [
                    {
                      "type": "Action.OpenUrl",
                      "title": "View Details",
                      "url": `https://aka.ms/spfx-adaptive-card-extensions-samples#${codeToUse}`
                    }
                  ]
                }
              ],
              "width": "stretch"
            }
          ]
        }
      ]
    };
  }
}
