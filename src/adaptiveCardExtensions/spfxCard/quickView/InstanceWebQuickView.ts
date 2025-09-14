import { BaseWebQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import {
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState
} from '../SpfxCardAdaptiveCardExtension';

export interface IQuickViewData {
  subTitle: string;
  title: string;
}

export class InstanceWebQuickView extends BaseWebQuickView<
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState
> {

  VIEW_ID = "UNKNOWN_VIEWID";

  constructor(viewId: string) {
    super();
    this.VIEW_ID = viewId;
  }
  render(): void {
    if (!this.state.SPFxExtensionInstance?.renderWebQuickView) {
      this.domElement.innerHTML = `
        <div>
          <h2>Error</h2>
          <p>App ${this.state.selectedApp} registered Web Quick View ${this.VIEW_ID}, but the instance does not have a renderWebQuickView method.</p>
        </div>
      `;
      return;
    }
    this.state.SPFxExtensionInstance?.renderWebQuickView(this.VIEW_ID, this.domElement);
  }
}
