import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFI } from '@pnp/sp';
import { SPFx as spSPFx } from '@pnp/sp/behaviors/spfx';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/folders';

import RecruitmentTracker from './components/RecruitmentTracker';
import { IRecruitmentTrackerProps } from './components/IRecruitmentTrackerProps';
import { GraphService } from './components/shared/GraphService';
import { ICurrentUser } from './components/shared/models';

export interface IRecruitmentTrackerWebPartProps {
  applyBaseUrl: string;
}

export default class RecruitmentTrackerWebPart extends BaseClientSideWebPart<IRecruitmentTrackerWebPartProps> {
  private _sp: SPFI = undefined!;
  private _graphService: GraphService = undefined!;
  private _currentUser: ICurrentUser = { displayName: 'Loading…', email: '' };

  protected async onInit(): Promise<void> {
    this._sp = spfi().using(spSPFx(this.context));
    this._graphService = new GraphService(this.context.msGraphClientFactory);

    try {
      this._currentUser = await this._graphService.getCurrentUser();
    } catch {
      this._currentUser = {
        displayName: this.context.pageContext.user.displayName,
        email: this.context.pageContext.user.email,
      };
    }
  }

  public render(): void {
    const element: React.ReactElement<IRecruitmentTrackerProps> = React.createElement(
      RecruitmentTracker,
      {
        sp: this._sp,
        graphService: this._graphService,
        currentUser: this._currentUser,
        applyBaseUrl: this.properties.applyBaseUrl ?? '',
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'OpRea — Recruitment Tracker Configuration' },
          groups: [
            {
              groupName: 'Candidate Application Form',
              groupFields: [
                PropertyPaneLabel('applyBaseUrl', {
                  text: 'The base URL of your Azure Static Web App. After deployment, paste the URL here. A unique apply link will be auto-generated for each job posted.',
                }),
                PropertyPaneTextField('applyBaseUrl', {
                  label: 'Apply Form Base URL',
                  placeholder: 'https://your-app.azurestaticapps.net',
                  description: 'Example: https://opats-apply.azurestaticapps.net',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
