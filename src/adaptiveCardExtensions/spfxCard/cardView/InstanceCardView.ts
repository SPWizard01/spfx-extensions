import {
    BaseComponentsCardView,
    ComponentsCardViewParameters,
    BasicCardView,
    BarChartCardView,
    LineChartCardView,
    PieChartCardView,
    ImageCardView,
    SearchCardView,
    TextInputCardView,
    TextInputImageCardView,
    PrimaryTextCardView,
    IExternalLinkCardAction,
    IQuickViewCardAction,
    IBasicTextCardViewConfiguration,
    IActionArguments,
    IPieChartCardViewConfiguration,
    IImageCardViewConfiguration,
    ISearchCardViewConfiguration,
    ITextInputCardViewConfiguration,
    ITextInputImageCardViewConfiguration,
    IPrimaryTextCardViewConfiguration
} from '@microsoft/sp-adaptive-card-extension-base';
import {
    ISpfxCardAdaptiveCardExtensionProps,
    ISpfxCardAdaptiveCardExtensionState,
} from '../SpfxCardAdaptiveCardExtension';
import { SPFxExtensionAppAdaptiveCardViewType } from '@spfx-extensions/core';


export class InstanceCardView extends BaseComponentsCardView<
    ISpfxCardAdaptiveCardExtensionProps,
    ISpfxCardAdaptiveCardExtensionState,
    ComponentsCardViewParameters
> {
    VIEW_ID = "UNKNOWN_VIEWID";
    CARD_TYPE: SPFxExtensionAppAdaptiveCardViewType | undefined;

    constructor(viewId: string, cardType: SPFxExtensionAppAdaptiveCardViewType) {
        super();
        this.VIEW_ID = viewId;
        this.CARD_TYPE = cardType;
    }

    getErrorCard(error: string, code: string): ComponentsCardViewParameters {
        return BasicCardView({
            cardBar: {
                componentName: "cardBar",
                title: "Error"
            },
            header: {
                componentName: "text",
                text: error
            },
            footer: {
                componentName: "cardButton",
                title: "View Details",
                action: {
                    type: "ExternalLink",
                    parameters: {
                        target: `https://aka.ms/spfx-adaptive-card-extensions-samples#${code}`
                    }
                },
                
            }
        })
    }

    onAction(action: IActionArguments): void {
        this.state.SPFxExtensionInstance?.onAction?.(this.VIEW_ID, action);
    }

    public get cardViewParameters(): ComponentsCardViewParameters {
        const params = this.state.SPFxExtensionInstance?.getCardViewParameters?.(this.VIEW_ID)
        if (!params) {
            return this.getErrorCard(`Selected app ${this.state.selectedApp} did not return data from getCardViewParameters for view ${this.VIEW_ID}`, "NO_CARDVIEW_PARAMETERS");
        }

        switch (this.CARD_TYPE) {
            case "BasicCardView":
                return BasicCardView(params as IBasicTextCardViewConfiguration);
            case "BarChartCardView":
                //eslint-disable-next-line @typescript-eslint/no-explicit-any
                return BarChartCardView(params as any);
            case "LineChartCardView":
                //eslint-disable-next-line @typescript-eslint/no-explicit-any
                return LineChartCardView(params as any);
            case "PieChartCardView":
                return PieChartCardView(params as IPieChartCardViewConfiguration);
            case "ImageCardView":
                return ImageCardView(params as IImageCardViewConfiguration);
            case "SearchCardView":
                return SearchCardView(params as ISearchCardViewConfiguration);
            case "TextInputCardView":
                return TextInputCardView(params as ITextInputCardViewConfiguration);
            case "TextInputImageCardView":
                return TextInputImageCardView(params as ITextInputImageCardViewConfiguration);
            case "PrimaryTextCardView":
                return PrimaryTextCardView(params as IPrimaryTextCardViewConfiguration);
        }

        return this.getErrorCard(`Selected app ${this.state.selectedApp} registered Card Type ${this.CARD_TYPE} for view ${this.VIEW_ID} which is not supported`, "NO_CARD_TYPE_FOUND");
    }

    public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
        return this.state.SPFxExtensionInstance?.onCardSelection?.(this.VIEW_ID);
    }
}
