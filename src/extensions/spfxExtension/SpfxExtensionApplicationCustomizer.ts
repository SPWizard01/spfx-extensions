import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
//import * as strings from 'SpfxExtensionApplicationCustomizerStrings';
import { initCore } from '../../services/initCoreService';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpfxExtensionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  _noProps: boolean;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxExtensionApplicationCustomizerProperties> {

  public async onInit(): Promise<void> {
    const envType =
      Environment.type === EnvironmentType.SharePoint
        ? "SharePoint"
        : "ClassicSharePoint";
    //init core then do stuff
    await initCore(
      this.context.pageContext,
      envType,
      this.context.placeholderProvider,
      this
    );
  }
}
