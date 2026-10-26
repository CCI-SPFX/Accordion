import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneLabel,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import Accordion from "./components/Accordion";
import { IAccordionProps } from "./components/IAccordionProps";

export interface IAccordionWebPartProps {
  listName: string;
  placeholderText: string;
  itemsToShow: number;
  searchBarEnabled: boolean;
  itemsHeight: number;
  itemsGap: number;

  // Button / search bar style
  buttonFontSize: number;
  buttonRadius: number;
  buttonFontWeight: string;
  buttonBackgroundColor: string;
  buttonTextColor: string;

  // Panel (answer / body) style
  panelFontSize: number;
  panelFontWeight: string;
  panelBackgroundColor: string;
  panelTextColor: string;
}

export default class AccordionWebPart extends BaseClientSideWebPart<IAccordionWebPartProps> {
  private listOptions: { key: string; text: string }[] = [];
  private listsLoaded: boolean = false;

  public render(): void {
    const element: React.ReactElement<IAccordionProps> = React.createElement(
      Accordion,
      {
        listName: this.properties.listName,
        placeholderText: this.properties.placeholderText,
        itemsToShow: this.properties.itemsToShow,
        itemsHeight: this.properties.itemsHeight,
        itemsGap: this.properties.itemsGap,
        searchBarEnabled: this.properties.searchBarEnabled,
        context: this.context,

        buttonFontSize: this.properties.buttonFontSize,
        buttonRadius: this.properties.buttonRadius,
        buttonFontWeight: this.properties.buttonFontWeight,
        buttonBackgroundColor: this.properties.buttonBackgroundColor,
        buttonTextColor: this.properties.buttonTextColor,

        panelFontSize: this.properties.panelFontSize,
        panelFontWeight: this.properties.panelFontWeight,
        panelBackgroundColor: this.properties.panelBackgroundColor,
        panelTextColor: this.properties.panelTextColor,
      },
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Sensible defaults so the web part looks fine before the user touches the property pane
    if (this.properties.itemsToShow === undefined) {
      this.properties.itemsToShow = 4;
    }

    if (!this.properties.buttonFontSize) {
      this.properties.buttonFontSize = 14;
    }
    if (this.properties.buttonRadius === undefined) {
      this.properties.buttonRadius = 6;
    }
    if (!this.properties.buttonFontWeight) {
      this.properties.buttonFontWeight = "600";
    }
    if (!this.properties.buttonBackgroundColor) {
      this.properties.buttonBackgroundColor = "#0078D4";
    }
    if (!this.properties.buttonTextColor) {
      this.properties.buttonTextColor = "#FFFFFF";
    }

    if (!this.properties.panelFontSize) {
      this.properties.panelFontSize = 14;
    }
    if (!this.properties.panelFontWeight) {
      this.properties.panelFontWeight = "400";
    }
    if (!this.properties.panelBackgroundColor) {
      this.properties.panelBackgroundColor = "#F3F2F1";
    }
    if (!this.properties.panelTextColor) {
      this.properties.panelTextColor = "#201F1E";
    }

    return super.onInit();
  }

  /** Pulls all visible lists on the site so the user can pick one from a dropdown in the property pane */
  private _loadLists(): Promise<void> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Title&$orderby=Title`;
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value: { Title: string }[] }) => {
        this.listOptions = (data.value || []).map((l) => ({
          key: l.Title,
          text: l.Title,
        }));
        this.listsLoaded = true;
      })
      .catch(() => {
        this.listOptions = [];
        this.listsLoaded = true;
      });
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this.listsLoaded) {
      return;
    }
    this._loadLists().then(() => {
      this.context.propertyPane.refresh();
    });
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Accordion Web Part Settings" },
          groups: [
            {
              groupFields: [
                this.listsLoaded ?
                  PropertyPaneDropdown("listName", {
                    label: "SharePoint List",
                    options: this.listOptions,
                    selectedKey: this.properties.listName,
                  })
                : PropertyPaneLabel("listNameLoading", {
                    text: "Loading lists...",
                  }),
                PropertyPaneLabel("listColumnsInfo", {
                  text: "Title, Description, OrderNumber (default: 0)",
                }),
                PropertyPaneSlider("itemsToShow", {
                  label: "Number of items to show",
                  min: 1,
                  max: 50,
                  step: 1,
                  value: this.properties.itemsToShow,
                }),
                PropertyPaneSlider("itemsHeight", {
                  label: "Item height (px)",
                  min: 40,
                  max: 120,
                  step: 1,
                  value: this.properties.itemsHeight,
                }),
                PropertyPaneSlider("itemsGap", {
                  label: "Gap between items (px)",
                  min: 5,
                  max: 25,
                  step: 1,
                  value: this.properties.itemsGap,
                }),
              ],
            },
            {
              groupName: "Button / Search Bar Style",
              groupFields: [
                PropertyPaneToggle("searchBarEnabled", {
                  label: "Enable Search Bar",
                  checked: this.properties.searchBarEnabled,
                }),
                PropertyPaneTextField("placeholderText", {
                  label: "Placeholder text for search bar",
                  value: this.properties.placeholderText,
                }),
                PropertyPaneSlider("buttonFontSize", {
                  label: "Font size (px)",
                  min: 10,
                  max: 36,
                  step: 1,
                  value: this.properties.buttonFontSize,
                }),
                PropertyPaneSlider("buttonRadius", {
                  label: "Corner radius (px)",
                  min: 0,
                  max: 40,
                  step: 1,
                  value: this.properties.buttonRadius,
                }),
                PropertyPaneDropdown("buttonFontWeight", {
                  label: "Font weight",
                  selectedKey: this.properties.buttonFontWeight,
                  options: [
                    { key: "400", text: "Normal" },
                    { key: "500", text: "Medium" },
                    { key: "600", text: "Semi Bold" },
                    { key: "700", text: "Bold" },
                    { key: "800", text: "Extra Bold" },
                  ],
                }),
                PropertyPaneTextField("buttonBackgroundColor", {
                  label: "Background color (hex, e.g. #0078D4)",
                  value: this.properties.buttonBackgroundColor,
                }),
                PropertyPaneTextField("buttonTextColor", {
                  label: "Text color (hex, e.g. #FFFFFF)",
                  value: this.properties.buttonTextColor,
                }),
              ],
            },
            {
              groupName: "Answer Panel Style",
              groupFields: [
                PropertyPaneSlider("panelFontSize", {
                  label: "Font size (px)",
                  min: 10,
                  max: 36,
                  step: 1,
                  value: this.properties.panelFontSize,
                }),
                PropertyPaneDropdown("panelFontWeight", {
                  label: "Font weight",
                  selectedKey: this.properties.panelFontWeight,
                  options: [
                    { key: "400", text: "Normal" },
                    { key: "500", text: "Medium" },
                    { key: "600", text: "Semi Bold" },
                    { key: "700", text: "Bold" },
                  ],
                }),
                PropertyPaneTextField("panelBackgroundColor", {
                  label: "Background color (hex, e.g. #F3F2F1)",
                  value: this.properties.panelBackgroundColor,
                }),
                PropertyPaneTextField("panelTextColor", {
                  label: "Text color (hex, e.g. #201F1E)",
                  value: this.properties.panelTextColor,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
