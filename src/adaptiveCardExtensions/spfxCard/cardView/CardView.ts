import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  BasicCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import {
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState,
  QUICK_VIEW_REGISTRY_ID
} from '../SpfxCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    return BasicCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: "SPFx Adaptive Card Extension"
      },
      footer: {
        componentName: 'cardButton',
        title: "Quick View",
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://www.bing.com'
      }
    };
  }
}
