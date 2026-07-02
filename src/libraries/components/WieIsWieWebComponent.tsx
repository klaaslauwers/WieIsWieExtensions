/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/explicit-function-return-type */

import * as React from "react";
import * as ReactDOM from "react-dom";

import { BaseWebComponent } from "@pnp/modern-search-extensibility";
import { Icon } from "@fluentui/react/lib/Icon";
import { Persona, PersonaPresence, PersonaSize } from "@fluentui/react/lib/components/Persona";

import { SPComponentLoader } from "@microsoft/sp-loader";
import { ServiceScope } from "@microsoft/sp-core-library";


import { IUserService } from "../services/IUserService";
import { UserService } from "../services/UserService";
import { ISearchResult } from "../data/ISearchResult";
import { IUserPresence } from "../data/IUserPresence";


import styles from "./WieIsWieWebComponent.module.scss";

const LIVE_PERSONA_COMPONENT_ID: string = "914330ee-2df2-4f6e-a858-30c23a812408";
const DEFAULT_FALLBACK_IMAGE_URL: string = "https://dendermonde.sharepoint.com/SiteAssets/Search/default-user.png";

export interface IObjectParam {
  searchResultItem?: ISearchResult;
  fallbackImage?: string;
}

export interface IWieIsWieWebComponentProps {
  searchResultItem?: ISearchResult;
  fallbackImage?: string;
  userService: IUserService;
  serviceScope: ServiceScope;
}

export interface IWieIsWieWebComponentState {
  user: ISearchResult;
  isComponentLoaded: boolean;
  presenceStatus: { [key: string]: any };
  currentPresence?: IUserPresence;
  isLoading: boolean;
}

export class WieIsWieComponent extends React.Component<
  IWieIsWieWebComponentProps,
  IWieIsWieWebComponentState
> {
  private sharedLibrary: any;
  private _isMounted: boolean = false;
  private _loadRequestId: number = 0;

  public constructor(props: IWieIsWieWebComponentProps) {
    super(props);

    this.sharedLibrary = null;

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
      currentPresence: undefined,
      isLoading: true,
      user: this._createEmptyUser()
    };

    this._loadSpfxSharedLibrary(); // eslint-disable-line @typescript-eslint/no-floating-promises
  }

  private get _fallbackImageUrl(): string {
    return this.props.fallbackImage || DEFAULT_FALLBACK_IMAGE_URL;
  }

  private _createEmptyUser(): ISearchResult {
    return {
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
      RefinableString119: undefined,
      PictureThumbnailURL: this._fallbackImageUrl
    } as any;
  }

  private _getSearchResultKey(item?: ISearchResult): string {
    if (!item) {
      return "";
    }

    return (
      item.UniqueID ||
      item.NormUniqueID ||
      item.Path ||
      item.ListItemID?.toString() ||
      item.DocId?.toString() ||
      ""
    );
  }

  private _getDisplayName(user: ISearchResult): string {
    const firstName = user.RefinableString101 || "";
    const lastName = user.RefinableString102 || "";
    const fullName = `${firstName} ${lastName}`.trim();

    return fullName || user.RefinableString09 || user.Title || "";
  }

  private _getEmail(user: ISearchResult): string {
    return ((user as any).Mail || user.RefinableString100 || "").toString().trim();
  }

  private _getJobTitle(user: ISearchResult): string {
    return ((user as any).JobTitle || user.RefinableString103 || "").toString().trim();
  }

  private _getPhotoUrl(user: ISearchResult): string {
    const email = this._getEmail(user);

    if (!email || !user.SPSiteURL) {
      return this._fallbackImageUrl;
    }

    return `${user.SPSiteURL}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`;
  }

  private async _loadSpfxSharedLibrary(): Promise<void> {
    if (this.state.isComponentLoaded) {
      return;
    }

    try {
      this.sharedLibrary = await SPComponentLoader.loadComponentById(LIVE_PERSONA_COMPONENT_ID);

      if (this._isMounted) {
        this.setState({
          isComponentLoaded: true
        });
      }
    } catch (error) {
      console.error("Failed to load LivePersonaCard shared library", error);
    }
  }

  private async _loadUser(searchResultItem?: ISearchResult): Promise<void> {
    const loadRequestId = ++this._loadRequestId;

    if (!searchResultItem) {
      if (this._isMounted) {
        this.setState({
          user: this._createEmptyUser(),
          currentPresence: undefined,
          isLoading: false
        });
      }

      return;
    }

    if (this._isMounted) {
      this.setState({
        user: this._createEmptyUser(),
        currentPresence: undefined,
        isLoading: true
      });
    }

    try {
      const user: ISearchResult = await this.props.userService.getUserData(
        searchResultItem,
        "Account"
      );

      if (!this._isMounted || loadRequestId !== this._loadRequestId) {
        return;
      }

      const photoUrl = this._getPhotoUrl(user);
      (user as any).PictureThumbnailURL = photoUrl || this._fallbackImageUrl;

      let currentPresence: IUserPresence | undefined = undefined;

      if (user.UserId) {
        try {
          currentPresence = await this.props.userService.getPresenceStatusCurrent(user.UserId);
        } catch (presenceError) {
          console.warn("Could not load presence for user", user.UserId, presenceError);
          currentPresence = undefined;
        }
      }

      if (!this._isMounted || loadRequestId !== this._loadRequestId) {
        return;
      }

      const currentKey = this._getSearchResultKey(this.props.searchResultItem);
      const loadedKey = this._getSearchResultKey(searchResultItem);

      if (currentKey !== loadedKey) {
        return;
      }

      this.setState({
        user,
        currentPresence,
        isLoading: false
      });
    } catch (error) {
      console.error("Failed to load Wie-is-Wie user", error);

      if (this._isMounted && loadRequestId === this._loadRequestId) {
        this.setState({
          user: this._createEmptyUser(),
          currentPresence: undefined,
          isLoading: false
        });
      }
    }
  }

  public async componentDidMount(): Promise<void> {
    this._isMounted = true;
    await this._loadUser(this.props.searchResultItem);
  }

  public async componentDidUpdate(prevProps: IWieIsWieWebComponentProps): Promise<void> {
    const previousKey = this._getSearchResultKey(prevProps.searchResultItem);
    const currentKey = this._getSearchResultKey(this.props.searchResultItem);

    if (previousKey !== currentKey) {
      await this._loadUser(this.props.searchResultItem);
    }
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    this._loadRequestId++;
  }

  public render(): React.ReactElement {
    const user = this.state.user;
    const email = this._getEmail(user);
    const displayName = this._getDisplayName(user);
    const jobTitle = this._getJobTitle(user);
    const imageUrl = ((user as any).PictureThumbnailURL || this._fallbackImageUrl).toString();

    const firstName = user.RefinableString101 || "";
    const lastName = user.RefinableString102 || "";

    const showTeamsActions = user.RefinableString117 !== "True" && !!email;

    const persona = (
      <Persona
        hidePersonaDetails={true}
        imageUrl={imageUrl}
        imageAlt={displayName}
        size={PersonaSize.size120}
        coinSize={150}
        presence={
          this.state.currentPresence && this.state.currentPresence.availability
            ? this.state.presenceStatus[this.state.currentPresence.availability]
            : PersonaPresence.none
        }
        presenceTitle={
          this.state.currentPresence && this.state.currentPresence.activity
            ? this.state.currentPresence.activity
            : "PresenceUnknown"
        }
      />
    );

    const livePersonaCard =
      this.state.isComponentLoaded && this.sharedLibrary && email ? (
        <this.sharedLibrary.LivePersonaCard
          className={`livePersonaCard ${styles.profileImage}`}
          clientScenario="PeopleWebPart"
          disableHover={false}
          hostAppPersonaInfo={{
            PersonaType: "User",
            PersonaAadObjectId: user.UserId,
            PersonaDisplayName: displayName
          }}
          serviceScope={this.props.serviceScope}
          upn={email}
          email={email}
          onCardOpen={() => console.log("LivePersonaCard Open")}
          onCardClose={() => console.log("LivePersonaCard Close")}
        >
          {persona}
        </this.sharedLibrary.LivePersonaCard>
      ) : (
        <div className={styles.profileImage}>
          {persona}
        </div>
      );

    return (
      <div className={styles.profile}>
        {livePersonaCard}

        <div className={styles.profileTop}>
          <div className={styles.profileName}>
            <span>
              <strong>
                {firstName} {lastName}
              </strong>
            </span>

            {showTeamsActions && (
              <span>
                <span className={styles.profilePropertyTeams}>
                  <a
                    href={`https://teams.microsoft.com/l/call/0/0?users=${encodeURIComponent(email)}`}
                    target="_blank"
                    rel="noreferrer"
                  >
                    <Icon iconName="Phone" aria-hidden="true" />
                  </a>
                </span>

                <span className={styles.profilePropertyTeams}>
                  <a
                    href={`https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(email)}`}
                    target="_blank"
                    rel="noreferrer"
                  >
                    <Icon iconName="CannedChat" aria-hidden="true" />
                  </a>
                </span>
              </span>
            )}
          </div>

          {jobTitle && (
            <div className={styles.profileTitle}>
              {jobTitle}
            </div>
          )}

          {user.RefinableString119 && (
            <div className={styles.profileProperty}>
              {user.RefinableString119}
            </div>
          )}

          {email && (
            <div className={styles.profileProperty}>
              <span>
                <Icon iconName="Mail" aria-hidden="true" />
                &nbsp;
                {email}
              </span>
            </div>
          )}

          {(user.RefinableString105 || user.RefinableString106) && (
            <div className={`${styles.profileProperty} ${styles.profilePropertyMargin}`}>
              {user.RefinableString105 && (
                <>
                  <span>
                    <Icon iconName="Phone" aria-hidden="true" />
                    &nbsp;
                    {user.RefinableString105}
                  </span>
                  {user.RefinableString106 && <span>|</span>}
                </>
              )}

              {user.RefinableString106 && (
                <>
                  <span>
                    <Icon iconName="CellPhone" aria-hidden="true" />
                    &nbsp;
                    {user.RefinableString106}
                  </span>
                </>
              )}
            </div>
          )}

          {user.RefinableString107 && (
            <div className={styles.profileProperty}>
              <span>
                <Icon iconName="MapPin" aria-hidden="true" />
                &nbsp;
                {user.RefinableString107}
              </span>
            </div>
          )}

          {(user.RefinableString108 || user.RefinableString109) && (
            <div className={`${styles.profileProperty} ${styles.profilePropertyMargin}`}>
              {user.RefinableString109 && (
                <>
                  <span>
                    <Icon iconName="Focus" aria-hidden="true" />
                    &nbsp;
                    {user.RefinableString109}
                  </span>
                  {user.RefinableString108 && <span>|</span>}
                </>
              )}

              {user.RefinableString108 && (
                <>
                  <span>
                    <Icon iconName="Communications" aria-hidden="true" />
                    &nbsp;
                    {user.RefinableString108}
                  </span>
                </>
              )}
            </div>
          )}
        </div>
      </div>
    );
  }
}

export class WieIsWieComponentWebComponent extends BaseWebComponent {
  private _userService: IUserService;

  public static get observedAttributes(): string[] {
    return [
      "data-search-result-item",
      "data-fallback-image"
    ];
  }

  public constructor() {
    super();
    this._userService = new UserService(this._serviceScope);
  }

  public connectedCallback(): void {
    this._renderComponent();
  }

  public attributeChangedCallback(
    attributeName: string,
    oldValue: string,
    newValue: string
  ): void {
    if (oldValue === newValue) {
      return;
    }

    if (this.isConnected) {
      this._renderComponent();
    }
  }

  public disconnectedCallback(): void {
    ReactDOM.unmountComponentAtNode(this);
  }

  private _renderComponent(): void {
    const props = this.resolveAttributes() as IObjectParam;

    const customComponent = (
      <WieIsWieComponent
        {...props}
        userService={this._userService}
        serviceScope={this._serviceScope}
      />
    );

    ReactDOM.render(customComponent, this);
  }
}