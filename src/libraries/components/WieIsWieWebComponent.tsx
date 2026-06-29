/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import { BaseWebComponent } from "@pnp/modern-search-extensibility";
import * as ReactDOM from "react-dom";
import { Icon } from '@fluentui/react/lib/Icon';
import { IUserService } from "../services/IUserService";
import { UserService } from "../services/UserService";
import styles from './WieIsWieWebComponent.module.scss';
import { ISearchResult } from "../data/ISearchResult";
import { Persona, PersonaPresence, PersonaSize } from "@fluentui/react/lib/components/Persona";
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ServiceScope } from '@microsoft/sp-core-library';
import { IUserPresence } from "../data/IUserPresence";

const LIVE_PERSONA_COMPONENT_ID: string = "914330ee-2df2-4f6e-a858-30c23a812408";

export interface IObjectParam {
  searchResultItem?: ISearchResult;
}

export interface IWieIsWieWebComponentProps {
  searchResultItem?: ISearchResult;
  userService: IUserService;
  serviceScope: ServiceScope;
}

export interface IWieIsWieWebComponentState {
  user: ISearchResult;
  isComponentLoaded: boolean;
  presenceStatus: { [key: string]: any };
  currentPresence?: IUserPresence;
}

export class WieIsWieComponent extends React.Component<
  IWieIsWieWebComponentProps,
  IWieIsWieWebComponentState
> {
  
  private sharedLibrary: any; // eslint-disable-line @typescript-eslint/no-explicit-any

  constructor(props: IWieIsWieWebComponentProps) {
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
    const personaCoin = <Persona
      hidePersonaDetails={true}
      imageUrl={this.state.user.PictureThumbnailURL}
      size={PersonaSize.size120}
      coinSize={150}
      presence={(this.state.currentPresence && this.state.currentPresence.availability) ? this.state.presenceStatus[this.state.currentPresence.availability] : PersonaPresence.none}
      presenceTitle={(this.state.currentPresence && this.state.currentPresence.activity) ? this.state.currentPresence.activity : "PresenceUnknown"}
    />;

    const lpcPersona = this.state.isComponentLoaded ? 
          React.createElement(
            this.sharedLibrary.LivePersonaCard,
            {
              className: `livePersonaCard ${styles.profileImage}`,
              clientScenario: "PeopleWebPart",
              disableHover: false,
              hostAppPersonaInfo: {
                PersonaType: "User",
                PersonaAadObjectId: this.state.user.UserId,
                PersonaDisplayName: this.state.user.RefinableString101 + " " + this.state.user.RefinableString102
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
    
    return (
      <div className={styles.profile}>
        {/* <img className={styles.profileImage} src={profileImageUrl} /> */}
        {lpcPersona}
        <div className={styles.profileTop}>
          <div className={styles.profileName}>
            <span><strong>{this.state.user.RefinableString101} {this.state.user.RefinableString102}</strong></span>
            {this.state.user.RefinableString117 !== "True" && (
              <span>
                <span className={styles.profilePropertyTeams}><a href={`https://teams.microsoft.com/l/call/0/0?users=${this.state.user.Mail}`} target="_blank"><Icon iconName="Phone" aria-hidden="true" /></a></span>
                <span className={styles.profilePropertyTeams}><a href={`https://teams.microsoft.com/l/chat/0/0?users=${this.state.user.Mail}`} target="_blank"><Icon iconName="CannedChat" aria-hidden="true" /></a></span>
              </span>
            )
          }
          </div>
          {this.state.user.RefinableString103 && <div className={styles.profileTitle}>{this.state.user.RefinableString103}</div>}
          {this.state.user.RefinableString119 && <div className={styles.profileProperty}>{this.state.user.RefinableString119}</div>}
          {this.state.user.Mail && <div className={styles.profileProperty}><span><Icon iconName="Mail" />&nbsp;{this.state.user.Mail}</span></div>}
          {(this.state.user.RefinableString105 || this.state.user.RefinableString106) &&
            <div className={`${styles.profileProperty} ${styles.profilePropertyMargin}`}>
              {this.state.user.RefinableString105 && 
              <>
                <span><Icon iconName="Phone" aria-hidden="true" />&nbsp;{this.state.user.RefinableString105}</span>
                <span>|</span>
              </>
              }
              {this.state.user.RefinableString106 &&
              <>
                <span><Icon iconName="CellPhone" aria-hidden="true" />&nbsp;{this.state.user.RefinableString106}</span>
              </>
              }
            </div>
          }
          {this.state.user.RefinableString107 && <div className={styles.profileProperty}><span><Icon iconName="MapPin" aria-hidden="true" />&nbsp;{this.state.user.RefinableString107}</span></div>
          }
          {(this.state.user.RefinableString108 || this.state.user.RefinableString109) &&
            <div className={`${styles.profileProperty} ${styles.profilePropertyMargin}`}>
              {this.state.user.RefinableString109 && 
              <>
                <span><Icon iconName="Focus" aria-hidden="true" />&nbsp;{this.state.user.RefinableString109}</span>
                <span>|</span>
              </>
              }
              {this.state.user.RefinableString108 && 
              <>
                <span><Icon iconName="Communications" aria-hidden="true" />&nbsp;{this.state.user.RefinableString108}</span>
              </>
              }
            </div>
          }
        </div>
      </div>
    );
  }
}

export class WieIsWieComponentWebComponent extends BaseWebComponent {
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
      <WieIsWieComponent {...props} userService={this._userService} serviceScope={this._serviceScope} />
    );
    ReactDOM.render(customComponent, this);
  }
}