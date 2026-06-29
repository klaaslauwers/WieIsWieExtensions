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
import { ISearchResult } from "../data/ISearchResult";
import { Persona, PersonaPresence, PersonaSize } from "@fluentui/react/lib/components/Persona";
import { IUserPresence } from "../data/IUserPresence";
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ServiceScope } from '@microsoft/sp-core-library';

const LIVE_PERSONA_COMPONENT_ID: string = "914330ee-2df2-4f6e-a858-30c23a812408";

export interface IObjectParam {
  searchResultItem?: ISearchResult;
}

export interface INieuwkomerWebComponentProps {
  searchResultItem?: ISearchResult;
  userService: IUserService;
  serviceScope: ServiceScope;
}

export interface INieuwkomerWebComponentState {
  user: ISearchResult;
  isComponentLoaded: boolean;
  presenceStatus: { [key: string]: any };
  currentPresence?: IUserPresence;
}

export class NieuwkomerComponent extends React.Component<
  INieuwkomerWebComponentProps,
  INieuwkomerWebComponentState
> {
  
  private sharedLibrary: any; // eslint-disable-line @typescript-eslint/no-explicit-any

  constructor(props: INieuwkomerWebComponentProps) {
    super(props);
    this.sharedLibrary = null;
    // Maping presence status MSGraph to persona presence status
    const presenceStatus = [] as { [key: string]: any };
    presenceStatus.Available = PersonaPresence.online;
    presenceStatus.AvailableIdle = PersonaPresence.online;
    presenceStatus.Away = PersonaPresence.away;
    presenceStatus.BeRightBack = PersonaPresence.away;
    presenceStatus.Busy = PersonaPresence.busy;
    presenceStatus.BusyIdle = PersonaPresence.busy;
    presenceStatus.DoNotDisturb = PersonaPresence.dnd;
    presenceStatus.Offline = PersonaPresence.offline;
    presenceStatus.PresenceUnknown = PersonaPresence.none;
    this.state = {
      isComponentLoaded: false,
      presenceStatus,
      user: {
        AuthorOWSUSER: "",
        contentclass: "",
        ContentTypeId: "",
        Created: new Date(),
        CreatedBy: "",
        DefaultEncodingURL: "",
        HtmlFileType: "",
        JobTitle: "",
        ListItemID: 0,
        NormListID: "",
        NormSiteID: "",
        NormUniqueID: "",
        NormWebID: "",
        Path: "",
        RefinableString09: "",
        RefinableDate03: new Date(),
        RefinableDate04: new Date(),
        SiteLogo: "",
        SiteTitle: "",
        SPSiteURL: "",
        SPWebUrl: "",
        Title: "",
        UniqueID: "",
        ResultTypeIdList: 0,
        ResultTypeId: 0,
        RenderTemplateId: "",
        piSearchResultId: "",
        DocId: 0,
        Rank: 0,
        GeoLocationSource: "",
        SiteId: "",
        WebId: "",
        IsExternalContent: false,
        ListId: "",
        UrlZone: 0,
        OriginalPath: "",
        IdentitySiteCollectionId: "",
        IdentityWebId: "",
        IdentityListId: "",
        IdentityListItemId: "",
        LastModifiedTime: new Date(),
        DocumentSignature: "",
        CollapsingStatus: 0,
        AutoPreviewUrl: "",
        key: 0,
        RefinableString100: undefined,
        RefinableString101: undefined,
        RefinableString102: undefined,
        RefinableString103: undefined,
        RefinableString104: undefined,
        RefinableString105: undefined,
        RefinableString106: undefined,
        RefinableString107: undefined,
        RefinableString108: undefined,
        UserId: "",
        RefinableString109: undefined,
        RefinableString116: undefined,
        RefinableString117: undefined,
        RefinableString119: undefined
      }
    };
    this._loadSpfxSharedLibrary(); // eslint-disable-line @typescript-eslint/no-floating-promises
  }

  private async _loadSpfxSharedLibrary(): Promise<void> {
    if (!this.state.isComponentLoaded) {
      try {
        this.sharedLibrary = await SPComponentLoader.loadComponentById(LIVE_PERSONA_COMPONENT_ID);

        this.setState({
          isComponentLoaded: true
        });

      } catch (error) {
        console.error(error);
      }
    }
  }

  //https://github.com/YannickRe/spfx-msgraph-peoplesearch/blob/main/src/components/PeopleViewComponent/PeopleViewComponent.tsx
  async componentDidMount(): Promise<void> {
    if (this.props.searchResultItem) {
      let user: ISearchResult = await this.props.userService.getUserData(
        this.props.searchResultItem,
        "Account"
      );

      const profileImage: string = await this.props.userService.getImageBase64FromUrl(user.SPSiteURL + "/_layouts/15/userphoto.aspx?size=L&username=" + user.Mail, user.UserId);
      user.PictureThumbnailURL = profileImage;
      const currentPresence: IUserPresence = await this.props.userService.getPresenceStatusCurrent(user.UserId);

      this.setState({ user, currentPresence });
    }
  }

  public render() {
    let jobTitle: string | undefined = this.state.user.JobTitle;
    if(!jobTitle || jobTitle === '') {
      jobTitle = this.state.user.RefinableString103;
    }

    console.log(jobTitle);

    const personaCoin = <Persona
      styles={{
        root: {},
        primaryText: {
          fontSize: 14,
          fontWeight: 500,
        },
      }}
      imageUrl={this.state.user.PictureThumbnailURL}
      text={`${this.state.user.RefinableString101} ${this.state.user.RefinableString102}`}
      secondaryText={jobTitle}
      size={PersonaSize.size48}
      coinSize={48}
      presence={(this.state.currentPresence && this.state.currentPresence.availability) ? this.state.presenceStatus[this.state.currentPresence.availability] : PersonaPresence.none}
      presenceTitle={(this.state.currentPresence && this.state.currentPresence.activity) ? this.state.currentPresence.activity : "PresenceUnknown"}
    />;

    return this.state.isComponentLoaded ? 
      React.createElement(
        this.sharedLibrary.LivePersonaCard,
        {
          className: 'livePersonaCard',
          clientScenario: "PeopleWebPart",
          disableHover: false,
          hostAppPersonaInfo: {
            PersonaType: "User",
            PersonaAadObjectId: this.state.user.UserId,
            PersonaDisplayName: this.state.user.RefinableString09
          },
          serviceScope: this.props.serviceScope,
          upn: this.state.user.Mail,
          email: this.state.user.Mail,
          onCardOpen: () => {
            console.log('LivePersonaCard Open');
          },
          onCardClose: () => {
            console.log('LivePersonaCard Close');
          },
        },
        personaCoin) : 
        personaCoin;
  }
}

export class NieuwkomerComponentWebComponent extends BaseWebComponent {
  private _userService: IUserService;
  public constructor() {
    super();
    this._userService = new UserService(this._serviceScope);
  }

  public async connectedCallback() {
    const props = this.resolveAttributes();
    const customComponent = (
      <NieuwkomerComponent {...props} userService={this._userService} serviceScope={this._serviceScope} />
    );
    ReactDOM.render(customComponent, this);
  }
}