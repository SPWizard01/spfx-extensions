import { DisplayMode, Environment, EnvironmentType, Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConditionalGroup,
  type IPropertyPaneConfiguration,
  IPropertyPaneCustomFieldProps,
  IPropertyPaneDropdownOption,
  IPropertyPaneDropdownProps,
  IPropertyPaneField,
  IPropertyPaneGroup,
  PropertyPaneButton,
  PropertyPaneDropdown,
  PropertyPaneFieldType,
  PropertyPaneLabel
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from "@microsoft/sp-webpart-base";
import type { IReadonlyTheme } from "@microsoft/sp-component-base";

import styles from "./SpfxExtensionloaderWebPart.module.scss";
//import * as strings from "SpfxExtensionloaderWebPartStrings";
import { SPFxExtensionAppConfig, SPFxExtensionAppDefinition, SPFxExtensionAppIcon, SPFxExtensionAppInstance, SPFxExtensionAppRuntimeConfig, SPFxExtensionAppSearchableData } from "@spfx-extensions/core";
import { APP_BUTTON_LABEL, EDIT_PAGE_AND_SELECT_WEBPART, SELECT_WEBPART, SPFXPREFIX } from "../../utilities/constants";

export interface ISpfxExtensionloaderWebPartProps extends SPFxExtensionAppSearchableData {
  selectedApp: string;
  SPFxExtensionAppConfiguration: SPFxExtensionAppConfig | undefined;
}

type propertyPath = keyof ISpfxExtensionloaderWebPartProps;



export default class SpfxExtensionloaderWebPart extends BaseClientSideWebPart<ISpfxExtensionloaderWebPartProps> {

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

  emptyRendered = false;
  _isDarkTheme: boolean = false;

  public async onInit() {
    const envType =
      Environment.type === EnvironmentType.SharePoint
        ? "SharePoint"
        : "ClassicSharePoint";
    const { initCore } = await import(/* webpackChunkName: "spfx-extension-loader" */"../../services/initCoreService");
    //init core then do stuff;
    await initCore(envType);
    this.appCatalogUrl = window.__SPFxExtensions.Utils.ConfiguratorPageUrl;
    if (this.properties.selectedApp) {
      this.mountApp(this.properties.selectedApp).catch((err) => {
        console.error(SPFXPREFIX, "Error while mounting appid", this.properties.selectedApp, err);
      });
    }
  }

  protected onPropertyPaneConfigurationComplete(): void {
    const isPaneOpen = this.context.propertyPane.isPropertyPaneOpen();

    // notify close only if the pane is not open
    // complete event fires also when config is saved
    //     This event method is invoked in the following cases:

    // When the CONFIGURATION_COMPLETE_TIMEOUT((currently the value is 5 secs) elapses after the last change.

    // When user clicks the "X" (close) button before the CONFIGURATION_COMPLETE_TIMEOUT elapses.

    // When user clicks the 'Apply' button before the CONFIGURATION_COMPLETE_TIMEOUT elapses.

    // When the user switches web parts then the current web part gets this event.
    if (!isPaneOpen && this.SPFxExtensionInstance) {

      this.SPFxExtensionInstance.executeListeners(
        "onConfigurationClose",
        undefined
      );
    }
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: propertyPath,
    //eslint-disable-next-line @typescript-eslint/no-explicit-any
    oldValue: any,
    //eslint-disable-next-line @typescript-eslint/no-explicit-any
    newValue: any
  ): void {
    // if selected app changed unmount the old app
    if (propertyPath === "selectedApp") {
      if (oldValue && oldValue !== newValue && this.SPFxExtensionInstance) {
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
        this.mountApp(newValue).catch((err) => {
          console.error(SPFXPREFIX, "Error while mounting appid", newValue, err);
        });
      }
    }
  }

  protected onDisplayModeChanged(oldDisplayMode: DisplayMode): void {
    const mode = oldDisplayMode === DisplayMode.Edit ? "Read" : "Edit";
    if (this.SPFxExtensionInstance) {
      this.SPFxExtensionInstance.executeListeners("onDisplayModeChange", mode);
      return;
    }
  }

  openPropertyPane() {
    // if (this.context.propertyPane.isPropertyPaneOpen()) {
    //   this.context.propertyPane.close();
    // }
    this.context.propertyPane.open();
  }

  closePropertyPane() {
    this.context.propertyPane.close();
  }

  isPropertyPaneOpen() {
    return this.context.propertyPane.isPropertyPaneOpen();
  }

  saveConfigValue(config: SPFxExtensionAppConfig, raiseEvent = true) {
    // const a = config.searchableText;
    // delete config.searchableText;
    // this.properties.searchableText = a;
    this.properties.SPFxExtensionAppConfiguration = config;
    if (raiseEvent) {
      this.SPFxExtensionInstance?.executeListeners("onConfigurationChange", config);
    }
  }

  getConfigValue(key?: string) {
    if (key) {
      return (this.properties[key as keyof ISpfxExtensionloaderWebPartProps] as SPFxExtensionAppConfig | undefined);
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

  private async mountApp(appId: string) {
    if (ISDEBUG) {
      console.debug(SPFXPREFIX, "Mounting app", appId, "at", this.domElement);
    }
    const runTimeConfig: SPFxExtensionAppRuntimeConfig = {
      domElement: this.domElement,
      //eslint-disable-next-line @typescript-eslint/no-explicit-any
      webpartContext: this.context as any,
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
    };
    this.SPFxExtensionInstance = await window.__SPFxExtensions.InstantiateApp(appId, runTimeConfig);
    if (!this.SPFxExtensionInstance) {
      return;
    }
    const newApp = window.__SPFxExtensions.Apps.find((app) => app.id === appId);
    if (newApp) {
      this.appDescription = newApp.description;
      //spfx specific, for some reason refresh does not work properly (custom field is not rerendered)
      // this.context.propertyPane.refresh();
      // if (this.context.propertyPane.isPropertyPaneOpen()) {
      //   this.context.propertyPane.close();
      //   this.context.propertyPane.open();
      // }
    }
  }

  private unmountApp() {
    if (this.SPFxExtensionInstance && this.SPFxExtensionInstance.unmount) {
      if (ISDEBUG) {
        console.debug(
          SPFXPREFIX,
          "Unmounting app",
          this.SPFxExtensionInstance.key,
          "at",
          this.SPFxExtensionInstance.domElement
        );
      }
      this.SPFxExtensionInstance.unmount();
    }

    this.properties.SPFxExtensionAppConfiguration = undefined;
  }

  protected onDispose(): void {
    this.unmountApp();
  }

  appButtonElements: HTMLElement[] = [];

  webpartSectionElement = document.createElement("section");
  webpartSectionTitle = document.createElement("header");
  appButtonsWrapper = document.createElement("div");
  appButtonsContainer = document.createElement("div");

  generateIconElement(icon?: SPFxExtensionAppIcon) {
    const iconElement = document.createElement("i");
    iconElement.classList.add(styles.icon);

    if (!icon) {
      iconElement.innerHTML = `<svg fill="currentColor" class="___12fm75w f1w7gpdv fez10in fg4l7m0" aria-hidden="true" width="28" height="28" viewBox="0 0 28 28" xmlns="http://www.w3.org/2000/svg"><path d="M20.84 2.66a2.25 2.25 0 0 0-3.18 0L13.5 6.8v-.56c0-1.24-1-2.25-2.25-2.25h-7C3.01 4 2 5.01 2 6.25v18c0 .97.78 1.75 1.75 1.75h18c1.24 0 2.25-1 2.25-2.25v-7c0-1.24-1-2.25-2.25-2.25h-.56l4.16-4.15c.88-.88.88-2.3 0-3.19l-4.5-4.5ZM17.31 14.5H13.5v-3.8l3.8 3.8Zm1.41-10.78c.3-.3.77-.3 1.06 0l4.5 4.5c.3.3.3.77 0 1.06l-4.5 4.51c-.3.3-.77.3-1.06 0l-4.5-4.5a.75.75 0 0 1 0-1.07l4.5-4.5ZM12 6.25v8.25H3.5V6.25c0-.41.34-.75.75-.75h7c.41 0 .75.34.75.75Zm-8.5 17.5V16H12v8.5H4.25a.75.75 0 0 1-.75-.75Zm10-7.75h8.25c.41 0 .75.34.75.75v7c0 .42-.34.75-.75.75H13.5V16Z" fill="currentColor"></path></svg>`;
      return iconElement;
    }

    if (icon.iconType === "font" && icon.fontFamily) {
      iconElement.style.fontFamily = icon.fontFamily;
      iconElement.classList.add(styles.iconFont);
    }

    if (icon.iconType === "url") {
      const imageElement = document.createElement("img");
      imageElement.src = icon.iconData;
      iconElement.appendChild(imageElement);
    }

    if (icon.iconType === "svg") {
      iconElement.innerHTML = icon.iconData;
    }

    return iconElement;
  }

  createAndAppendAppButtons(app: SPFxExtensionAppDefinition) {
    const appButtonElement = document.createElement("button");
    appButtonElement.title = APP_BUTTON_LABEL;
    appButtonElement.ariaLabel = APP_BUTTON_LABEL;
    appButtonElement.className = styles.appButton;

    const icon = this.generateIconElement(app.icon);
    appButtonElement.append(icon, app.name);
    appButtonElement.title = app.name;

    appButtonElement.addEventListener("click", () => {
      this.properties.selectedApp = app.id;
      this.webpartSectionElement.remove();
      this.mountApp(app.id).catch((err) => {
        console.error(SPFXPREFIX, "Error while mounting app", app, err);
      });
    });

    this.appButtonsContainer.appendChild(appButtonElement);
  }

  createWebpartSection(button?: boolean) {
    this.webpartSectionElement.className = styles.applicationListSection;
    this.webpartSectionTitle.className = styles.header;
    this.webpartSectionElement.appendChild(this.webpartSectionTitle);
    if (button) {
      this.webpartSectionElement.appendChild(this.appButtonsWrapper);
    }
  }

  renderDisplayMode() {
    this.webpartSectionElement.ariaLabel = EDIT_PAGE_AND_SELECT_WEBPART;
    this.webpartSectionTitle.textContent = EDIT_PAGE_AND_SELECT_WEBPART;
    this.createWebpartSection();
    this.domElement.appendChild(this.webpartSectionElement);
  }

  renderEditMode() {
    this.webpartSectionElement.ariaLabel = SELECT_WEBPART;
    this.webpartSectionTitle.textContent = SELECT_WEBPART;
    this.createWebpartSection(true);

    this.appButtonsContainer.className = styles.appButtonsContainer;
    this.appButtonsWrapper.appendChild(this.appButtonsContainer);
    this.appButtonsWrapper.className = styles.appButtonsWrapper;
    this.domElement.appendChild(this.webpartSectionElement);

    window.__SPFxExtensions.AddAppEventListener("appAdded", (app) => {
      if (app.isWebPartApp) {
        this.createAndAppendAppButtons(app);
      }
    });

    window.__SPFxExtensions.Apps.filter((app) => app.registrationCompleted).forEach(
      (app) => {
        if (app.isWebPartApp) {
          this.createAndAppendAppButtons(app);
        }
      }
    );

    window.__SPFxExtensions.Utils.spAppInitializationPromise.then(() => {
      this.domElement.appendChild(this.webpartSectionElement);

      window.__SPFxExtensions.Utils.appManifestPromises.forEach((promise) => {
        const buttonLoader = document.createElement("div");
        const loaderSpinner = document.createElement("span");
        buttonLoader.className = styles.buttonLoader;
        loaderSpinner.className = styles.loader;
        buttonLoader.appendChild(loaderSpinner);

        this.appButtonsContainer.append(buttonLoader);
        promise
          .catch(() => {
            //do nothing
          }).finally(() => {
            buttonLoader.remove();
          });
      });
    }).catch(() => {
      // do nothing
    });
  }

  public render(): void {
    // do not uncomment this or this will "erase the webpart when exiting edit mode";
    // required when adding same Webpart while another instance is already open and configuration pane is open as well.
    if (this.context.propertyPane.isPropertyPaneOpen()) {
      this.onPropertyPaneConfigurationStart();
    }
    if (this.properties.selectedApp) {
      return;
    }

    this.domElement.className = styles.SPFxExtensionApp;

    if (this.displayMode === DisplayMode.Read) {
      this.renderDisplayMode();
    }

    if (this.displayMode === DisplayMode.Edit) {
      this.renderEditMode();
    }
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      // selectedApp: {
      //   isSearchablePlainText: true,
      // },
      // SPFxExtensionAppConfiguration: {
      //   dynamicPropertyType: "object",
      // },
      searchableText: {
        isSearchablePlainText: true,
      },
      searchableHtml: {
        isHtmlString: true,
      },
    };
  }

  protected onPropertyPaneConfigurationStart(): void {
    // wait for all the manifests to load
    window.__SPFxExtensions.AllAppAssetsLoadedPromise.then(() => {
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
        (app) => app.isWebPartApp
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
      // refresh to rerender the dropdown and description
      this.context.propertyPane.refresh();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    }).catch((err: any) => {
      console.error(SPFXPREFIX, "Error while awaiting all app assets to load", err);
    });
  }

  CustomWebpartConfigurationField(
    name: string
  ): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
    return {
      type: PropertyPaneFieldType.Custom,
      targetProperty: name,
      properties: {
        key: "SPFxExtensionAppConfiguration",
        onRender: (domElement, _context, _callBack) => {
          this.configDomElement = domElement;
          if (this.SPFxExtensionInstance) {
            // when app instance is loaded forward the render event
            this.SPFxExtensionInstance.instanceLoadPromise.then(() => {
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
          }
        },
        onDispose: (_domeElement, _context) => {
          if (this.SPFxExtensionInstance) {
            this.SPFxExtensionInstance.executeListeners(
              "onConfigurationClose",
              undefined
            );
          }
        },
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
            window.open(`${this.appCatalogUrl}?web=${this.context.pageContext.web.absoluteUrl}`, "_blank");
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
                this.CustomWebpartConfigurationField(
                  "SPFxExtensionAppConfiguration"
                ),
              ],
            },
          ],
        },
      ],
    };
  }



  // private _getEnvironmentMessage(): Promise<string> {
  //   if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
  //     return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
  //       .then(context => {
  //         let environmentMessage: string = "";
  //         switch (context.app.host.name) {
  //           case "Office": // running in Office
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
  //             break;
  //           case "Outlook": // running in Outlook
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
  //             break;
  //           case "Teams": // running in Teams
  //           case "TeamsModern":
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
  //             break;
  //           default:
  //             environmentMessage = strings.UnknownEnvironment;
  //         }

  //         return environmentMessage;
  //       });
  //   }

  //   return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  // }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty("--bodyText", semanticColors.bodyText || null);
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty("--linkHovered", semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }


}
