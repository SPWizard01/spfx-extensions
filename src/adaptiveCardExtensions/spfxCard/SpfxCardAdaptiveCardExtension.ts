//#region Imports
import {
  IPropertyPaneConditionalGroup,
  IPropertyPaneCustomFieldProps,
  IPropertyPaneDropdownOption,
  IPropertyPaneDropdownProps,
  IPropertyPaneField,
  IPropertyPaneGroup,
  PropertyPaneButton,
  PropertyPaneDropdown,
  PropertyPaneFieldType,
  PropertyPaneLabel,
  type IPropertyPaneConfiguration
} from "@microsoft/sp-property-pane";
import {
  AdaptiveCardExtensionContext,
  BaseAdaptiveCardExtension,
  IAdaptiveCardExtensionPropertiesMetadata,
  IOnBeforeActionArguments,
  RenderType
} from "@microsoft/sp-adaptive-card-extension-base";
import {
  SPFxExtensionAppAdaptiveCardInstance,
  SPFxExtensionAppAdaptiveCardRuntimeConfig,
  SPFxExtensionAppConfig,
  SPFxExtensionAppSearchableData,
  SPFxExtensionCleanup,
} from "@spfx-extensions/core";
import { NoAppCardView } from "./cardView/NoAppCardView";
import { SPFXPREFIX } from "../../utilities/constants";
import { ThemeProvider } from "@microsoft/sp-component-base";
import { DisplayMode, ServiceScope } from '@microsoft/sp-core-library';
import { ERROR_CARD_VIEW_REGISTRY_ID, ErrorCardView } from "./cardView/ErrorCardView";
import { GetRegistrators } from "./SpfxCardRegistrations";
// import * as strings from 'SpfxCardAdaptiveCardExtensionStrings';
//#endregion

export interface ISpfxCardAdaptiveCardExtensionProps extends SPFxExtensionAppSearchableData {
  selectedApp: string;
  SPFxExtensionAppConfiguration: SPFxExtensionAppConfig | undefined;
}

export interface ISpfxCardAdaptiveCardExtensionState {
  SPFxExtensionInstance: SPFxExtensionAppAdaptiveCardInstance | undefined;
  selectedApp: string;
  error: string;
  errorCode: string;
}
type propertyPath = keyof ISpfxCardAdaptiveCardExtensionProps;

const NO_APP_CARD_VIEW: string = 'SPFX_NO_APP_CARD_VIEW';

export default class SpfxCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ISpfxCardAdaptiveCardExtensionProps,
  ISpfxCardAdaptiveCardExtensionState
> {

  configuratorUrl = "/sites/appcatalog/SPFxExtensionsData/SitePages/SPFxExtensionsConfigurator.aspx";
  coreInitPromise = new Promise((resolve) => {
    import(/* webpackChunkName: "spfx-extension-loader" */"../../services/initCoreService").then(({ initCore }) => {
      initCore("SharePoint").then(() => {
        this.configuratorUrl = window.__SPFxExtensions.Utils.ConfiguratorPageUrl;
        resolve(true);
      }).catch((e) => {
        console.error(SPFXPREFIX, "Initializing SPFxExtensions Core from WebPart failed", e);
      })
    }).catch((e) => {
      console.error(SPFXPREFIX, "Importing SPFxExtensions Core from WebPart failed", e);
    })
  });


  // SPFxExtensionInstance: SPFxExtensionAppInstance | undefined;
  allApps: IPropertyPaneDropdownOption[] = [];
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

  // for some reason onRender these properties are not available if accessing `this` on edit mode
  // so we copy them in onInit
  webPartContext!: AdaptiveCardExtensionContext;
  webPartComponentId!: string;
  webPartInstanceId!: string;

  initCleanup: SPFxExtensionCleanup | undefined;

  async PopulateDropdown() {
    try {
      await window.__SPFxExtensions.AllAppAssetsLoadedPromise;
      // register description if an app is matching this webpart
      const selectedApp = window.__SPFxExtensions.Apps.find(
        (app) => app.id === this.properties.selectedApp
      );
      if (selectedApp) {
        this.appDescription = selectedApp.description;
        this.hideAppSelectorWhenAppLoaded =
          selectedApp.hideAppSelectorWhenAppLoaded ?? false;
        this.hideConfiguratorButton = selectedApp.hideConfiguratorButton ?? false;
      }

      // Clear dropdown options in propertypane
      this.dropDownProps.options?.splice(0, this.dropDownProps.options?.length);

      const appOptionsInDropdown: IPropertyPaneDropdownOption[] = window.__SPFxExtensions.Apps.filter(
        (app) => app.instanceType === "adaptiveCard" && !app.hideWebPartButton
      ).map((app) => {
        return {
          key: app.id,
          text: app.name,
        };
      });

      this.dropDownProps.options?.push(...appOptionsInDropdown);

      // select key
      this.dropDownProps.selectedKey = this.properties.selectedApp;
      // enable dropdown
      this.dropDownProps.disabled = false;
    }
    catch (err) {
      console.error(SPFXPREFIX, "Error while awaiting all app assets to load", err);
    }
  }

  public async onInit() {
    this.state = {
      SPFxExtensionInstance: undefined,
      selectedApp: this.properties.selectedApp,
      error: "",
      errorCode: ""
    };

    if (DEBUG) {
      console.debug(SPFXPREFIX, "onInitCardExtension", this.instanceId);
    }

    this.themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this.serviceScope = this.context.serviceScope;

    this.webPartContext = this.context;
    this.webPartInstanceId = this.instanceId;
    this.webPartComponentId = this.componentId;
    this.cardNavigator.register(ERROR_CARD_VIEW_REGISTRY_ID, () => new ErrorCardView());
    this.cardNavigator.register(NO_APP_CARD_VIEW, () => new NoAppCardView());
    await this.coreInitPromise;
    let appInstance: SPFxExtensionAppAdaptiveCardInstance | undefined;
    if (this.properties.selectedApp && !this.state.SPFxExtensionInstance) {
      appInstance = await this.mountApp(this.properties.selectedApp);
    }
    if (appInstance) {
      await appInstance.instanceLoadPromise;
      this.setState({
        SPFxExtensionInstance: appInstance,
      });
      if (!appInstance.registerViews) {
        this.setState({
          error: `App ${this.properties.selectedApp} instance ${appInstance.key} does not implement registerViews method, cannot proceed.`,
          errorCode: "NO_REGISTER_VIEWS"
        })
        return;
      }
      if (DEBUG) {
        console.debug(SPFXPREFIX, "Registering views for app instance", appInstance.key);
      }
      await appInstance.registerViews(GetRegistrators(this.cardNavigator, this.quickViewNavigator));
      //should return after this but for debugging we continue
    } else {
      this.setState({
        error: `App ${this.properties.selectedApp} instance has been requested but core could not instantiate it.`,
        errorCode: "APP_INSTANCE_NOTRETURNED"
      })
      return;
    }
    this.initCleanup = window.__SPFxExtensions.AddAppEventListener("appAdded", (app) => {
      if (app.instanceType === "adaptiveCard" && !app.hideWebPartButton) {
        if (!this.dropDownProps.options?.some((o) => o.key === app.id)) {
          this.dropDownProps.options?.push({ key: app.id, text: app.name });
        }
      }
    });

  }

  protected renderCard() {
    if (this.state.SPFxExtensionInstance) {
      const renderView = this.state.SPFxExtensionInstance.renderCard?.();
      if (!renderView) {
        if (!this.state.error) {
          const msg = `No card view returned from app instance ${this.instanceId}`;
          this.setState({
            error: msg,
            errorCode: "NO_CARD_VIEW_RETURNED"
          });
        }

        return ERROR_CARD_VIEW_REGISTRY_ID;
      }
      return renderView;
    }
    return NO_APP_CARD_VIEW;
  }


  //#region Event forwarding
  onBeforeAction(action: IOnBeforeActionArguments): void {
    if (DEBUG) {
      console.debug(SPFXPREFIX, "onBeforeAction", action);
    }
    this.state.SPFxExtensionInstance?.onBeforeAction?.(this.navigator.currentId ?? "", action);
  }

  onDispose(): void {
    if (DEBUG) {
      console.debug(SPFXPREFIX, "onDispose");
    }
    this.initCleanup?.();
    this.state.SPFxExtensionInstance?.onDispose?.();
    this.unmountApp();
  }

  protected onDisplayModeChanged(oldDisplayMode: DisplayMode): void {
    const newDisplayMode = oldDisplayMode === DisplayMode.Edit ? "Read" : "Edit";
    if (DEBUG) {
      console.debug(SPFXPREFIX, "onDisplayModeChanged", newDisplayMode);
    }
    this.state.SPFxExtensionInstance?.onDisplayModeChanged?.(newDisplayMode);
  }

  protected onRenderTypeChanged(oldRenderType: RenderType): void {
    const newRenderType = oldRenderType === "Card" ? "QuickView" : "Card";
    if (DEBUG) {
      console.debug(SPFXPREFIX, "onRenderTypeChanged", newRenderType);
    }
    this.state.SPFxExtensionInstance?.onRenderTypeChanged?.(newRenderType);
  }
  //#endregion

  //#region Property Pane
  protected async loadPropertyPaneResources(): Promise<void> {
    if (DEBUG) {
      console.debug(SPFXPREFIX, "loadPropertyPaneResources");
    }
    await this.PopulateDropdown();

    if (this.state.SPFxExtensionInstance?.loadPropertyPaneResources) {
      await this.state.SPFxExtensionInstance.loadPropertyPaneResources();
    }
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: propertyPath,
    //eslint-disable-next-line @typescript-eslint/no-explicit-any
    oldValue: any,
    //eslint-disable-next-line @typescript-eslint/no-explicit-any
    newValue: any
  ): void {
    if (DEBUG) {
      console.debug(SPFXPREFIX, "onPropertyPaneFieldChanged", propertyPath, oldValue, newValue);
    }
    // if selected app changed unmount the old app
    if (propertyPath === "selectedApp") {
      if (oldValue && oldValue !== newValue && this.state.SPFxExtensionInstance) {
        const shouldUnmount = confirm(
          "You are about to switch app, this will erase all previous app configuration. Are you sure?"
        );
        if (!shouldUnmount) {
          this.properties[propertyPath] = oldValue;
          return;
        }
        this.unmountApp();
      }
      // if new app was selected, mount it
      if (newValue) {
        this.webpartSectionElement.remove();
        this.mountApp(newValue).catch(() => {
          // do nothing
        });
      }
    }
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
          this.state.SPFxExtensionInstance?.instanceLoadPromise
            .then(() => {
              this.state.SPFxExtensionInstance?.executeListeners(
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
          this.state.SPFxExtensionInstance?.executeListeners("onConfigurationClose", { domElement });
          this.configDomElement = undefined;
        },
        // context: this.context,
      },
    };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const configuratorButton: IPropertyPaneGroup | IPropertyPaneConditionalGroup = {
      groupFields: [
        PropertyPaneLabel("spfxExtensionLoaderLabel", {
          text: `App not working? Try refreshing the page. Or go to the configuration page.`,
        }),
        PropertyPaneButton("configuratorButton", {
          text: "Open Configurator",
          buttonType: 1,
          onClick: () => {
            window.open(`${this.configuratorUrl}?web=${this.context.pageContext.web.absoluteUrl}`, "_blank");
          }
        })
      ]
    };
    const cfgButtonGroup = this.hideConfiguratorButton ? [] : [configuratorButton];

    const appSelector: IPropertyPaneGroup | IPropertyPaneConditionalGroup = {
      groupFields: [
        PropertyPaneDropdown("selectedApp", {
          label: "App",
          disabled: this.dropDownProps.disabled,
          options: this.dropDownProps.options,
          selectedKey: this.dropDownProps.selectedKey,
        }),
        PropertyPaneLabel("selectedAppDecription", {
          text: this.appDescription,
        }),
      ],
    }
    const cfgAppSelector = this.hideAppSelectorWhenAppLoaded ? [] : [appSelector];

    return {
      pages: [
        {
          groups: [
            ...cfgButtonGroup,
            ...cfgAppSelector,
            {
              groupFields: [
                this.CustomCardConfigurationField(
                  "SPFxExtensionAppConfiguration"
                ),
              ],
            },
          ],
        },
      ],
    };
  }
  //#endregion

  //#region App mounting and property forwarding
  openPropertyPane() {
    // if (this.context.propertyPane.isPropertyPaneOpen()) {
    //   this.context.propertyPane.close();
    // }
    this.context.propertyPane?.open();
  }

  closePropertyPane() {
    this.context.propertyPane?.close();
  }

  isPropertyPaneOpen() {
    return this.context.propertyPane?.isPropertyPaneOpen() ?? false;
  }

  saveConfigValue(config: SPFxExtensionAppConfig, raiseEvent = true) {
    // const a = config.searchableText;
    // delete config.searchableText;
    // this.properties.searchableText = a;
    this.properties.SPFxExtensionAppConfiguration = config;
    if (raiseEvent) {
      this.state.SPFxExtensionInstance?.executeListeners("onConfigurationChange", config);
    }
  }

  getConfigValue(key?: string) {
    if (key) {
      let dataByKey = (this.properties[key as keyof ISpfxCardAdaptiveCardExtensionProps] as SPFxExtensionAppConfig | undefined);
      if (typeof dataByKey === "undefined") {
        dataByKey = this.properties.SPFxExtensionAppConfiguration;
      }
      return dataByKey;
    }
    return this.properties.SPFxExtensionAppConfiguration;
  }

  getSearchData() {
    return {
      searchableText: this.properties.searchableText,
      searchableHtml: this.properties.searchableHtml,
    };
  }

  setSearchData(data: SPFxExtensionAppSearchableData) {
    this.properties.searchableText = data.searchableText;
    this.properties.searchableHtml = data.searchableHtml;
  }

  getThemeProvider() {
    return this.themeProvider;
  }

  getConfigDomElement() {
    return this.configDomElement;
  }

  getContext() {
    return this.webPartContext;
  }

  getServiceScope() {
    return this.serviceScope;
  }

  getDisplayMode() {
    const currentMode = this.displayMode === DisplayMode.Edit ? "Edit" : "Read";
    return currentMode;
  }
  //#endregion

  private async mountApp(appId: string) {
    if (ISDEBUG) {
      console.debug(SPFXPREFIX, "Mounting Adaptive Card Extension", appId);
    }
    //clean HTML
    try {
      const runTimeConfig: SPFxExtensionAppAdaptiveCardRuntimeConfig = {
        instanceType: "adaptiveCard",
        //eslint-disable-next-line @typescript-eslint/no-explicit-any
        cardExtension: this as any,
        openPropertyPane: () => {
          this.openPropertyPane();
        },
        closePropertyPane: () => {
          this.closePropertyPane();
        },
        isPropertyPaneOpen: () => {
          return this.isPropertyPaneOpen();
        },
        saveConfigValue: (config: SPFxExtensionAppConfig, raise = true) => {
          this.saveConfigValue(config, raise);
        },
        getConfigValue: (key?: string) => {
          return this.getConfigValue(key);
        },
        getSearchableData: () => {
          return this.getSearchData();
        },
        setSearchableData: (data: SPFxExtensionAppSearchableData) => {
          this.setSearchData(data);
        },
        getThemeProvider: () => {
          return this.getThemeProvider();
        },
        getConfigDomElement: () => {
          return this.getConfigDomElement();
        },
        getContext: () => {
          return this.getContext();
        },
        getServiceScope: () => {
          return this.getServiceScope();
        },
        setState: (newState) => {
          this.setState(newState);
        },
        cardSize: () => {
          return this.cardSize;
        },
        renderType: () => {
          return this.renderType;
        },
        displayMode: () => {
          return this.getDisplayMode();
        },
        cardNavigator: () => {
          return this.cardNavigator;
        },
        quickViewNavigator: () => {
          return this.quickViewNavigator;
        }

      };

      const appInstance = await window.__SPFxExtensions.InstantiateApp(appId, runTimeConfig) as SPFxExtensionAppAdaptiveCardInstance;
      if (!appInstance) {
        console.error(SPFXPREFIX, "App instance is undefined, cannot mount app", appId);
        return;
      }
      if (DEBUG) {
        console.debug(SPFXPREFIX, "App instance mounted", appInstance.registerViews);
      }
      const newApp = window.__SPFxExtensions.Apps.find((app) => app.id === appId);
      if (newApp) {
        this.appDescription = newApp.description;
      }
      return appInstance;
    }
    catch (err) {
      console.error(SPFXPREFIX, "Error while mounting appid", appId, err);
    }
  }

  private unmountApp() {
    if (this.state.SPFxExtensionInstance) {
      if (DEBUG) {
        console.debug(
          SPFXPREFIX,
          "Unmounting Adaptive Card Extension",
          this.state.SPFxExtensionInstance.key
        );
      }
      this.state.SPFxExtensionInstance.unmount?.();
    }
  }

  protected get propertiesMetadata(): IAdaptiveCardExtensionPropertiesMetadata {
    return {
      searchableText: {
        isSearchablePlainText: true,
      },
      searchableHtml: {
        isHtmlString: true,
      },
    }
  }
}
