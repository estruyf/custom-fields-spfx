import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneHorizontalRule
} from '@microsoft/sp-client-preview';

import { PropertyPaneLoggingField } from './PropertyPaneControls/PropertyPaneLoggingField';

import styles from './CustomFields.module.scss';
import * as strings from 'customFieldsStrings';
import { ICustomFieldsWebPartProps } from './ICustomFieldsWebPartProps';

export default class CustomFieldsWebPart extends BaseClientSideWebPart<ICustomFieldsWebPartProps> {
  loggingValue: any = {url: 'https://thisistheurl.com', response: {body:'test',status:{code:200, msg: 'OK'}}, message: 'This is just a test message', date: new Date().toUTCString()};

  public constructor(context: IWebPartContext) {
    super(context);

    this._updateLogging = this._updateLogging.bind(this);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.customFields}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">${this.properties.description}</p>
              <a href="https://github.com/SharePoint/sp-dev-docs/wiki" class="ms-Button ${styles.button}">
                <span class="ms-Button-label">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  private _updateLogging(): any {
    this.loggingValue["date"] = new Date().toUTCString();
    return this.loggingValue;
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLoggingField({
                  label: "Logging field",
                  description: "Logging field description",
                  value: this.loggingValue,
                  retrieve: this._updateLogging
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
