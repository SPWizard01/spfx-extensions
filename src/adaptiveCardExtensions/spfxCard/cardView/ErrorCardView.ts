import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  PrimaryTextCardView,
} from '@microsoft/sp-adaptive-card-extension-base';
import {
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState,

} from '../SpfxCardAdaptiveCardExtension';

export const ERROR_CARD_VIEW_REGISTRY_ID = "SPFX_ERROR_CARD_VIEW";
export class ErrorCardView extends BaseComponentsCardView<
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {

  public get cardViewParameters(): ComponentsCardViewParameters {
    return PrimaryTextCardView({
      cardBar: {
        componentName: 'cardBar',
        title: "Error",
      },
      header: {
        componentName: 'text',
        text: "Error In Card View"
      },
      body: {
        componentName: "text",
        text: `
        ${this.state.error}
        `
      },
      footer: {
        componentName: 'cardButton',
        title: "View Details",
        action: {
          type: 'ExternalLink',
          parameters: {
            target: `https://aka.ms/spfx-adaptive-card-extensions-samples#${this.state.errorCode}`
          }
        }
      }
    });
  }
}
