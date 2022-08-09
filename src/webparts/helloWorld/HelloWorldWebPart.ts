import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

import * as pnp from 'sp-pnp-js';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div>
      <table>
      <tr>
        <td>Please Enter the Software ID</td>
        <td><input type='text' id='txtID'/></td>
        <td><input type='submit' id='btnRead' value='Read Details'/></td>
      </tr>
      <tr>
        <td>Software Title</td>
        <td><input type='text' id='txtSoftwareTitle'/></td>
      </tr>
      <tr>
        <td>Software Name</td>
        <td><input type='text' id='txtSoftwareName'/></td>
      </tr>
      <tr>
        <td>Software Vendor</td>
        <td>
          <select id='ddSoftwareVendor'>
            <option value='Microsoft'>Microsoft</option>
            <option value='Sun'>Sun</option>
            <option value='Google'>Google</option>
            <option value='Oracle'>Oracle</option>
          </select>
          </td>
      </tr>

      <tr>
        <td>Software Version</td>
        <td><input type='text' id='txtSoftwareVersion'/></td>
      </tr>

      <tr>
        <td>Software Description</td>
        <td><textarea rows='5' cols='40' id='txtSoftwareDescription'> </textarea> </td>
      </tr>

      <tr>
        <td colspan='2' align='center'>
          <input type='submit' value='Insert Item' id='btnSubmit' />
          <input type='submit' value='Update' id='btnUpdate' />
          <input type='submit' value='Delete' id='btnDelete' />
          <input type='submit' value='Show All Records' id='btnReadAll' />
          <input type='submit' value='CLEAR' id='btnClear' />
        </td>
      </tr>

    </table>

    <div id='divStatus'></div>

    <div id='spListData'></div>
      </div>
    </section>`;
    this._bindEvents();
  }

  private _bindEvents():void {
    this.domElement.querySelector('#btnSubmit').addEventListener("click", ()=> {this._addListItem();});
  }

  private _addListItem():void {
    const softwareTitle: string = this.domElement.querySelector('#txtSoftwareTitle')['value'];
    const softwareName: string = this.domElement.querySelector('#txtSoftwareName')['value'];
    const softwareVendor: string = this.domElement.querySelector('#ddSoftwareVendor')['value'];
    const softwareVersion: string = this.domElement.querySelector('#txtSoftwareVersion')['value'];
    const softwareDescription: string = this.domElement.querySelector('#txtSoftwareDescription')['value'];

    const siteUrl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('SoftwareCatalog')/items";

    pnp.sp.web.lists.getByTitle('SoftwareCatalog').items.add({
      Title: softwareTitle,
      SoftwareVendor: softwareVendor,
      SoftwareName: softwareName,
      SoftwareVersion: softwareVersion,
      SoftwareDescription: softwareDescription,

    }).then (r=> {
      alert("Succes");
    })

  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
