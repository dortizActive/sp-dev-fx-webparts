import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'GraphCalendarWebPartStrings';
import GraphCalendar from './components/GraphCalendar';
import { IGraphCalendarProps } from './components/IGraphCalendarProps';
import * as microsoftTeams from '@microsoft/teams-js';
import { initializeIcons } from 'office-ui-fabric-react';
import { PropertyPaneSlider, PropertyPaneCheckbox, IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';

export interface IGraphCalendarWebPartProps {
  limit: number;
  showRecurrence: boolean;
  groupId: string;
}

export default class GraphCalendarWebPart extends BaseClientSideWebPart<IGraphCalendarWebPartProps> {
  private _teamsContext: microsoftTeams.Context;

  public render(): void {
    const element: React.ReactElement<IGraphCalendarProps> = React.createElement(
      GraphCalendar,
      {
        limit: this.properties.limit,
        showRecurrence: this.properties.showRecurrence,
        context: this.context,
        teamsContext: this._teamsContext,
        groupId: this.properties.groupId
        
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<any> {
    // create a new promise
    return new Promise<void>((resolve, _reject) => {

      // Sets a default if limit has not been defined
      if (this.properties.limit === undefined) {
        this.properties.limit = 100;
      }

      if (this.properties.showRecurrence === undefined) {
        this.properties.showRecurrence = true;
      }

      if(this.properties.groupId === undefined){
        this.properties.groupId = "58b3389f-de06-460e-85cf-93ae1ccaf8da";
      }

      // Sets the Teams context if in Teams
      if (this.context.sdks.microsoftTeams) {
        this._teamsContext = this.context.sdks.microsoftTeams.context;
        
        // Initialize the OUIF icons if in Teams
        initializeIcons();
          
        // resolve the promise
        resolve(undefined);
      } else {
        // resolve the promise
        resolve(undefined);
      }
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

 
/*   protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
 */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneSlider('limit', {
                  label: strings.EventsPerView,
                  max: 500,
                  min: 50
                }),
                PropertyPaneCheckbox('showRecurrence', {
                  text: strings.ShowRecurringEvents,
                  checked: true
                })
              ]
            },{
              groupName: 'Opciones de Calendario',
              groupFields: [
                PropertyPaneTextField('groupId', {
                  label: 'ID del Grupo',
                  placeholder: 'Ingrese el ID del Grupo',
                  
                })
              ]
            }
          ]
        } 
      ]
    };
  }
}
