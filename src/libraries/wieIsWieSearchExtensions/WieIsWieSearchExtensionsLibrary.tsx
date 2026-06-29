import { IAdaptiveCardAction, IComponentDefinition, IDataSourceDefinition, IExtensibilityLibrary, ILayoutDefinition, IQueryModifierDefinition, ISuggestionProviderDefinition } from "@pnp/modern-search-extensibility";
import * as Handlebars from "handlebars";
import { WieIsWieComponentWebComponent } from "../components/WieIsWieWebComponent";
import { NieuwkomerComponentWebComponent } from "../components/NieuwkomerWebComponent";
import { MailButtonComponentWebComponent } from "../components/MailButtonWebComponent";
import { ContactDetailsComponentWebComponent } from "../components/ContactDetailsWebComponent";

export class WieIsWieSearchExtensionsLibrary implements IExtensibilityLibrary {

  public registerHandlebarsCustomizations(namespace: typeof Handlebars) {

    namespace.registerHelper('encodeUriComponent', (uriComp: string) => {
      if (uriComp) {
        return uriComp.replace('+', '%2B');
      }
      return '';
    });

    namespace.registerHelper('createMails', (items: string) => {
      if (items) {
        const searchResults: any = JSON.parse(items);
        const arrayOfUsers: string[] = searchResults.map((x: any) => x.RefinableString100).filter((x: string) => !!x);
        return arrayOfUsers.join(',');
      }
      return '';
    });

    namespace.registerHelper('notcontains', (stringToTest: string, substring: string) => {
      return stringToTest.toLowerCase().indexOf(substring.toLowerCase()) < 0;
    });
  }

  getCustomLayouts(): ILayoutDefinition[] {
    return [];
  }

  getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: 'mail-button-component',
        componentClass: MailButtonComponentWebComponent
      },
      {
        componentName: 'wie-is-wie-component',
        componentClass: WieIsWieComponentWebComponent
      },
      {
        componentName: 'nieuwkomer-component',
        componentClass: NieuwkomerComponentWebComponent
      },
      {
        componentName: 'contact-details-component',
        componentClass: ContactDetailsComponentWebComponent
      }
    ];
  }

  getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }
  invokeCardAction(action: IAdaptiveCardAction): void {
    return;
  }
  getCustomQueryModifiers?(): IQueryModifierDefinition[] {
    return [];
  }
  getCustomDataSources?(): IDataSourceDefinition[] {
    return [];
  }
  public name(): string {
    return 'WieIsWieSearchExtensionsLibrary';
  }
}
