import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "GraphWebPartStrings";
import GraphTest from "./components/Graph";
import { IGraphProps } from "./components/IGraphProps";
import { ClientMode } from "./components/ClientMode";

export interface IGraphWebPartProps {
  clientMode: ClientMode;
}

export default class GraphWebPart extends BaseClientSideWebPart<IGraphWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IGraphProps> =
      React.createElement(GraphTest, {
        clientMode: this.properties.clientMode,
        context: this.context,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneChoiceGroup("clientMode", {
                  label: strings.ClientModeLabel,
                  options: [
                    { key: ClientMode.aad, text: "AadHttpClient" },
                    { key: ClientMode.graph, text: "MSGraphClient" },
                  ],
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
