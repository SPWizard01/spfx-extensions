import { IPropertyPaneCustomFieldProps, IPropertyPaneDropdownOption, IPropertyPaneDropdownProps, IPropertyPaneField, PropertyPaneFieldType, PropertyPaneTextField, type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { AdaptiveCardExtensionContext, BaseAdaptiveCardExtension, RenderType } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { SPFXPREFIX } from '../../utilities/constants';
import { ThemeProvider } from '@microsoft/sp-component-base';
import { SPFxExtensionAppInstance } from '@spfx-extensions/core';
import { ServiceScope } from '@microsoft/sp-core-library';
// import * as strings from 'SpfxCardAdaptiveCardExtensionStrings';

export interface ISpfxCardAdaptiveCardExtensionProps {
  title: string;
}

export interface ISpfxCardAdaptiveCardExtensionState {
}

const CARD_VIEW_REGISTRY_ID: string = 'SpfxCard_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'SpfxCard_QUICK_VIEW';

export default class SpfxCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState
> {

  SPFxExtensionInstance: SPFxExtensionAppInstance | undefined;
  allApps: IPropertyPaneDropdownOption[] = [];
  appCatalogUrl = "/sites/appcatalog";
  dropDownProps: Partial<IPropertyPaneDropdownProps> = {
    options: [],
    selectedKey: "",
    disabled: true,
  };
  appDescription = "";
  hideAppSelectorWhenAppLoaded = false;
  hideConfiguratorButton = false;
  configDomElement: HTMLElement | undefined;
  themeProvider: ThemeProvider | undefined;
  serviceScope: ServiceScope | undefined;
  appButtonElements: HTMLElement[] = [];

  webpartSectionElement = document.createElement("section");
  webpartSectionTitle = document.createElement("header");
  appButtonsWrapper = document.createElement("div");
  appButtonsContainer = document.createElement("div");
  // token/registration returned by AddAppEventListener so we can remove it
  private appAddedListenerRegistration: unknown | undefined;

  // for some reason onRender these properties are not available if accessing `this` on edit mode
  // so we copy them in onInit
  webPartContext!: AdaptiveCardExtensionContext;
  webPartComponentId!: string;
  webPartInstanceId!: string;

  public async onInit() {
    this.state = {};

    if (DEBUG) {
      console.debug(SPFXPREFIX, "onInit", this.instanceId);
    }

    this.themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this.serviceScope = this.context.serviceScope;

    this.webPartContext = this.context;
    this.webPartInstanceId = this.instanceId;
    this.webPartComponentId = this.componentId;

    const { initCore } = await import(/* webpackChunkName: "spfx-extension-loader" */"../../services/initCoreService");
    //init core then do stuff;
    await initCore("SharePoint");
    this.appCatalogUrl = window.__SPFxExtensions.Utils.ConfiguratorPageUrl;


    // registers the card view to be shown in a dashboard
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    // registers the quick view to open via QuickView action
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());
    // this.navigator.register
  }

  protected async loadPropertyPaneResources(): Promise<void> {
    // this.setState
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  public get iconProperty() {
    return "";
  }

  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Write 1-3 sentences describing the functionality of this component." },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: "Title"
                })
              ]
            }
          ]
        },
        {
          header: { description: "Write 1-3 sentences describing the functionality of this component." },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: "Title"
                })
              ]
            }
          ]
        }
      ]
    };
  }
  protected onRenderTypeChanged(oldRenderType: RenderType): void {

  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {

  }

  CustomCardConfigurationField(
    name: string
  ): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
    return {
      type: PropertyPaneFieldType.Custom,
      targetProperty: name,
      properties: {
        key: name,
        onRender: (domElement, _context, _callBack) => {
          this.configDomElement = domElement;
          // when app instance is loaded forward the render event
          this.SPFxExtensionInstance?.instanceLoadPromise
            .then(() => {
              this.SPFxExtensionInstance?.executeListeners(
                "onConfigurationRender",
                {
                  domElement,
                }
              );
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
            }).catch((err: any) => {
              console.error(SPFXPREFIX, "Error while awaiting app to load", err);
            });
        },
        onDispose: (domElement, _context) => {
          this.SPFxExtensionInstance?.executeListeners("onConfigurationClose", { domElement });
          this.configDomElement = undefined;
        },
        // context: this.context,
      },
    };
  }
}
