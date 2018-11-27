import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DecisionsDashboardWebPart.module.scss';
import * as strings from 'DecisionsDashboardWebPartStrings';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { sp, Web } from '@pnp/sp';

require('../../../node_modules/@fortawesome/fontawesome-free/css/all.css');

export interface IDecisionsDashboardWebPartProps {
  description: string;
  listName: string;
  inProgressViewURL: string;
  notApprovedViewURL: string;
  approvedViewURL: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export interface ISPListField {
  Title: string;
  InternalName: string;
}

export interface ISPListFields {
  value: ISPListField[];
}

export default class DecisionsDashboardWebPart extends BaseClientSideWebPart<IDecisionsDashboardWebPartProps> {

  public fields: ISPListFields;

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.decisionsDashboard }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">CSC Decision Records Dashboard</span>
              <p class="${ styles.subTitle }">Click on any of the status below to go to a detailed view</p>
            </div>
          </div>
          <div id="spListContainer" />
        </div>
      </div>`;

      //use Promise chaining to ensure the status are processed in the order we expect
      this._getListFields()
        .then(() => {
          return this._renderByStatusAsync(this.fields, "In Progress");
        })
        .then(() => {
          return this._renderByStatusAsync(this.fields, "Not Approved");
        })
        .then(() => {
          return this._renderByStatusAsync(this.fields, "Approved");
        });
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
                }),
                PropertyPaneTextField('listName', {
                  label: 'Decision Records List Name'
                }),                
                PropertyPaneTextField('inProgressViewURL', {
                  label: 'URL to view for In Progress decisions'
                }),
                PropertyPaneTextField('notApprovedViewURL', {
                  label: 'URL to view for Not Approved decisions'
                }),
                PropertyPaneTextField('approvedViewURL', {
                  label: 'URL to view for Approved decisions'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _getListFields(): Promise<ISPListFields> {
    let web = new Web(this.context.pageContext.site.absoluteUrl);
web.get();
    return web.lists.getByTitle(this.properties.listName).fields.filter("Title eq 'Decision Status'").get();
      .then( fields => {

        var allFields: ISPListFields = { value: [] };

        for (var i = 0; i < fields.length; i++) {
          
          let currentField = { Title: fields[i].Title,
                               InternalName: fields[i].InternalName };

          allFields.value.push(currentField);
        }

        return allFields;
      });

  }

  private _renderByStatusAsync(fields: ISPListFields, status: string): Promise<any> {
    return this._getListDataByStatus(fields, status)
        .then((response) => {
          this._renderList(response.value, status);
        });
  }

  private _getListDataByStatus(fields: ISPListFields, status: string): Promise<ISPLists> {
    //we need to get the internal name of the fields for this query
    let decisionStatus: string = this._getFieldInternalName(fields, "Decision Status");

    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.properties.listName}')/items?$filter=${ decisionStatus } eq '` + status + `'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getFieldInternalName(fields: ISPListFields, fieldDisplayName: string): string {
    let fieldName: string = "";

    for (var i = 0; i < fields.value.length; i++) {
      if (fields.value[i].Title == fieldDisplayName)
        fieldName = fields.value[i].InternalName;
    }

    return fieldName;
  }

  private _renderList(items: ISPList[], status: string): void {
    let html: string = '';
    let count: number = items == undefined ? 0 : items.length;
    let style: string = '';
    let viewURL: string = '#';

    switch (status) {
      case "In Progress":
        style = styles.inProgress;
        viewURL = this.properties.inProgressViewURL == undefined ? "#" : this.properties.inProgressViewURL;
        break;
      case "Not Approved":
        style = styles.notApproved;
        viewURL = this.properties.notApprovedViewURL == undefined ? "#" : this.properties.notApprovedViewURL;
        break;
      case "Approved":
        style = styles.approved;
        viewURL = this.properties.approvedViewURL == undefined ? "#" : this.properties.approvedViewURL;
        break;
    }

    html += `
    <ul class="${styles.list}">
      <li class="${styles.listItem}">
        <a href='${viewURL}'><span class="ms-font-l ${style} fa wide">${status} Tasks: ${count}</span></a>
      </li>
    </ul>`;
 
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML += html;
  }

}
