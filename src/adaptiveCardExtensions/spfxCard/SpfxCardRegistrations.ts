import { BaseCardView, BaseQuickView, IQuickViewNavigator, IViewNavigator } from "@microsoft/sp-adaptive-card-extension-base";
import { InstanceCardView } from "./cardView/InstanceCardView";
import { InstanceQuickView } from "./quickView/InstanceQuickView";
import { InstanceWebQuickView } from "./quickView/InstanceWebQuickView";
import { SPFxExtensionAppAdaptiveCardRegistrators, SPFxExtensionAppAdaptiveCardViewType } from "@spfx-extensions/core";


export function GetRegistrators(cardNavigator: IViewNavigator<BaseCardView>, quickViewNavigator: IQuickViewNavigator<BaseQuickView>): SPFxExtensionAppAdaptiveCardRegistrators {
    function RegisterViewCard(uniqueViewId: string, cardType: SPFxExtensionAppAdaptiveCardViewType) {
        cardNavigator.register(uniqueViewId, () => new InstanceCardView(uniqueViewId, cardType));
    }
    function RegisterQuickView(uniqueViewId: string) {
        quickViewNavigator.register(uniqueViewId, () => new InstanceQuickView(uniqueViewId));
    }
    function RegisterWebQuickView(uniqueViewId: string) {
        quickViewNavigator.register(uniqueViewId, () => new InstanceWebQuickView(uniqueViewId));
    }

    return {
        RegisterViewCard,
        RegisterQuickView,
        RegisterWebQuickView
    };
}