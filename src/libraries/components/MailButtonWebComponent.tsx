/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import { BaseWebComponent } from "@pnp/modern-search-extensibility";
import * as ReactDOM from "react-dom";
import { IUserService } from "../services/IUserService";
import { UserService } from "../services/UserService";
import styles from './MailButtonWebComponent.module.scss';
import { ISearchResult } from "../data/ISearchResult";
import { ServiceScope } from '@microsoft/sp-core-library';

export interface IObjectParam {
  searchResultItem?: ISearchResult;
}

export interface IMailButtonWebComponentProps {
  searchResultItems?: ISearchResult[];
  userService: IUserService;
  serviceScope: ServiceScope;
}

export interface IMailButtonWebComponentState {
  users: ISearchResult[];
}

export class MailButtonComponent extends React.Component<
  IMailButtonWebComponentProps,
  IMailButtonWebComponentState
> {
  constructor(props: IMailButtonWebComponentProps) {
    super(props);

    this.state = {
      users: []
    };
  }

  async componentDidMount(): Promise<void> {
    if (this.props.searchResultItems) {
      let newSearchResultItems: ISearchResult[] = [...this.props.searchResultItems];
      for (let i: number =0; i< this.props.searchResultItems.length; i++) {
        const addedProps: ISearchResult = await this.props.userService.getUserData(this.props.searchResultItems[i], "Account");
        newSearchResultItems[i].Mail =  addedProps.Mail;
      }
      this.setState({users: newSearchResultItems});
    }
  }

  public render() {
    return (
      <div data-automation-id="button-container">
          <div className="ms-TooltipHost ms-TooltipHostShim">
            <div className={styles.mailButton} data-is-visible="true">
              <a role="link" data-automationid="splitbuttonprimary"
                data-automation-id="button-web-part" target="_self" id="button_webpart"
                className="fui-Button r1alrhcs ms-Button root-166 ms-Button--primary ms-ButtonShim--primary" data-is-focusable="true"
                href={`mailto:?bcc=${this.state.users.map(x => x.Mail).join(';')}`}>
                <span className="ms-Button-flexContainer ms-ButtonShim-flexContainer">
                  <span className="ms-Button-textContainer ms-ButtonShim-textContainer">
                    <span className="ms-Button-label ms-ButtonShim-label">Mail alle personen</span>
                  </span>
                </span>
              </a>
            </div>
          </div>
        </div>);
  }
}

export class MailButtonComponentWebComponent extends BaseWebComponent {
  private _userService: IUserService;
  public constructor() {
    super();
    this._userService = new UserService(this._serviceScope);
  }

  // private _dataService: IDataService;
  // private _sp: SPFI;

  public async connectedCallback() {
    const props = this.resolveAttributes();
    const customComponent = (
      <MailButtonComponent {...props} userService={this._userService} serviceScope={this._serviceScope} />
    );
    ReactDOM.render(customComponent, this);
  }
}