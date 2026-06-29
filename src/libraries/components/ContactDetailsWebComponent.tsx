/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import { BaseWebComponent } from "@pnp/modern-search-extensibility";
import * as ReactDOM from "react-dom";
import { ServiceScope } from '@microsoft/sp-core-library';
import { ISearchResult } from "../data/ISearchResult";
import { ISharePointService } from "../services/ISharePointService";
import { SharePointService } from "../services/SharePointService";
import { Spinner, SpinnerSize } from "@fluentui/react";

export interface IObjectParam {
  searchResultItem?: ISearchResult;
}

export interface IContactDetailsComponentWebComponentProps {
  searchResultItem?: ISearchResult;
  sharepointService: ISharePointService;
  serviceScope: ServiceScope;
}

export interface IContactDetailsWebComponentState {
  contactDetails: string[];
  loaded: boolean;
}

export class ContactDetailsComponent extends React.Component<
  IContactDetailsComponentWebComponentProps,
  IContactDetailsWebComponentState
> {
  
  constructor(props: IContactDetailsComponentWebComponentProps) {
    super(props);
    this.state = {
      loaded: false,
      contactDetails: []
    };
  }
  async componentDidMount(): Promise<void> {
    if (this.props.searchResultItem) {
      const contacts: string[] = await this.props.sharepointService.getContactLookup(
        this.props.searchResultItem
      );
      const contactDetails = await Promise.all(
        contacts.map(async (contact) => {
          return await this.props.sharepointService.getContactData(contact);
        })
      );

      this.setState({ contactDetails, loaded: true });
    }
  }

  public render() {    
    const htmlContent = this.state.contactDetails
      .map(detail => `<span>${detail}</span>`)
      .join('<br/>');

    const element: React.ReactElement = React.createElement('div', {
      dangerouslySetInnerHTML: { __html: htmlContent }
    });

    return this.state.loaded ? 
      element : <Spinner label="Contactgegevens aan het laden..." size={SpinnerSize.large} />
  }
}

export class ContactDetailsComponentWebComponent extends BaseWebComponent {
  private _sharepointService: ISharePointService;
  public constructor() {
    super();
    this._sharepointService = new SharePointService(this._serviceScope);
  }

  public async connectedCallback() {
    const props = this.resolveAttributes();
    const customComponent = (
      <ContactDetailsComponent {...props} sharepointService={this._sharepointService} serviceScope={this._serviceScope} />
    );
    ReactDOM.render(customComponent, this);
  }
}