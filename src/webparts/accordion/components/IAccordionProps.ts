import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IAccordionProps {
  listName: string;
  placeholderText: string;
  itemsToShow: number;
  itemsHeight: number;
  itemsGap: number;
  searchBarEnabled: boolean;
  context: WebPartContext;

  // Button / search bar style
  buttonFontSize: number;
  buttonRadius: number;
  buttonFontWeight: string;
  buttonBackgroundColor: string;
  buttonTextColor: string;

  // Answer panel style
  panelFontSize: number;
  panelFontWeight: string;
  panelBackgroundColor: string;
  panelTextColor: string;
}