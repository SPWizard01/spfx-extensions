import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import {
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState
} from '../SpfxCardAdaptiveCardExtension';

export interface IQuickViewData {
  subTitle: string;
  title: string;
}

export class QuickView extends BaseAdaptiveCardQuickView<
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      subTitle: "Quick view",
      title: "SPFx Adaptive Card Extension"
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}
