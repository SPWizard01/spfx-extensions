import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  PrimaryTextCardView,
} from '@microsoft/sp-adaptive-card-extension-base';
import {
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState,
} from '../SpfxCardAdaptiveCardExtension';

export class NoAppCardView extends BaseComponentsCardView<
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {

  public get cardViewParameters(): ComponentsCardViewParameters {
    return PrimaryTextCardView({
      cardBar: {
        componentName: 'cardBar',
        title: "Available Apps",
      },
      header: {
        componentName: 'text',
        text: "No SPFx Extension Card Selected"
      },
      body: {
        componentName: "text",
        text: "Please enter edit mode and select a Card in the property pane to display here."
      },
      
      footer: {
        componentName: "cardButton",
        title: "Help & Guides",
        action: {
          type: 'ExternalLink',
          parameters: {
            target: 'https://aka.ms/spfx-adaptive-card-extensions-samples'
          }
        }
      }
    });
  }
}
