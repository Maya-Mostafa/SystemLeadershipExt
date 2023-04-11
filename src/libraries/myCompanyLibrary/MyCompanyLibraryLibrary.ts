import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import {  IExtensibilityLibrary, 
          IComponentDefinition, 
          ISuggestionProviderDefinition, 
          ISuggestionProvider,
          ILayoutDefinition, 
          LayoutType, 
          ILayout,
          IAdaptiveCardAction,
          LayoutRenderType,
          IQueryModifierDefinition,
          IQueryModifier,
          IDataSourceDefinition,
          IDataSource
} from "@pnp/modern-search-extensibility";
import * as Handlebars from "handlebars";
import { MyCustomComponentWebComponent } from "../CustomComponent";
import { Customlayout } from "../CustomLayout";
import { CustomSuggestionProvider } from "../CustomSuggestionProvider";
import { CustomQueryModifier } from "../CustomQueryModifier";
import { CustomDataSource } from "../CustomDataSource";

export class MyCompanyLibraryLibrary implements IExtensibilityLibrary {
  

  public static readonly serviceKey: ServiceKey<MyCompanyLibraryLibrary> =
  ServiceKey.create<MyCompanyLibraryLibrary>('SPFx:MyCustomLibraryComponent', MyCompanyLibraryLibrary);

  private _spHttpClient: SPHttpClient;
  private _pageContext: PageContext;
  private _currentWebUrl: string;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
      this._currentWebUrl = this._pageContext.web.absoluteUrl;
    });
  }

  public getCustomLayouts(): ILayoutDefinition[] {
    return [
      {
        name: 'PnP Custom layout (Handlebars)',
        iconName: 'Color',
        key: 'CustomLayoutHandlebars',
        type: LayoutType.Results,
        renderType: LayoutRenderType.Handlebars,
        templateContent: require('../custom-layout.html'),
        serviceKey: ServiceKey.create<ILayout>('PnP:CustomLayoutHandlebars', Customlayout),
      },
      {
        name: 'PnP Custom layout (Adaptive Cards)',
        iconName: 'Color',
        key: 'CustomLayoutAdaptive',
        type: LayoutType.Results,
        renderType: LayoutRenderType.AdaptiveCards,
        templateContent: JSON.stringify(require('../custom-layout.json'), null, "\t"),
        serviceKey: ServiceKey.create<ILayout>('PnP:CustomLayoutAdaptive', Customlayout),
      }
    ];
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: 'my-custom-component',
        componentClass: MyCustomComponentWebComponent
      }
    ];
  }

  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [
        {
          name: 'Custom Suggestions Provider',
          key: 'CustomSuggestionsProvider',
          description: 'A demo custom suggestions provider from the extensibility library',
          serviceKey: ServiceKey.create<ISuggestionProvider>('MyCompany:CustomSuggestionsProvider', CustomSuggestionProvider)
      }
    ];
  }

  public registerHandlebarsCustomizations(namespace: typeof Handlebars) {

    // Register custom Handlebars helpers
    // Usage {{myHelper 'value'}}
    namespace.registerHelper('myHelper', (value: string) => {
      return new namespace.SafeString(value.toUpperCase());
    });
  }

  public invokeCardAction(action: any): void {
    
    // Process the action based on type
    if (action.type == "Action.OpenUrl") {
      //window.open(action.url, "_blank");
      console.log("action.url", action.url);
    } else if (action.type == "Action.Submit") {
      // Process the action based on title
      switch (action.title) {

        case 'Click on item':
          // Invoke the currentUser endpoint
          this._spHttpClient.get(
          `${this._currentWebUrl}/_api/web/currentUser`,
          SPHttpClient.configurations.v1, 
          null).then((response: SPHttpClientResponse) => {
            response.json().then((json) => {
              console.log(JSON.stringify(json));
              
            });
          });
        break;
        
        case 'Add to my read list':
          console.log("action.data.itemId", action.data.itemId);
          const itemId = action.data.itemId.toString();
          let userData = {
            'accountName': "i:0#.f|membership|" + this._pageContext.user.email,
            'propertyName': "PDSBSystemLinks",
            'propertyValues': [itemId]
          },
          spOptions: ISPHttpClientOptions = {
              headers:{
                  "Accept": "application/json;odata=nometadata", 
                  "Content-Type": "application/json;odata=nometadata",
                  "odata-version": "",
              },
              body: JSON.stringify(userData)
          };
          this._spHttpClient.post(`https://pdsb1.sharepoint.com/_api/SP.UserProfiles.PeopleManager/SetMultiValuedProfileProperty`,  SPHttpClient.configurations.v1, spOptions);
          break;

        case 'Global click':
          alert(JSON.stringify(action.data));
          break;

        default:
          console.log('Action not supported!');
          break;
      }
    }
  }

  public getCustomQueryModifiers(): IQueryModifierDefinition[]
  {
    return [
      {
        name: 'Word Modifier',
        key: 'WordModifier',
        description: 'A demo query modifier from the extensibility library',
        serviceKey: ServiceKey.create<IQueryModifier>('MyCompany:CustomQueryModifier', CustomQueryModifier)

      }
    ];
  
    }
  public getCustomDataSources(): IDataSourceDefinition[] {
    return [
      {
          name: 'NPM Search',
          iconName: 'Database',
          key: 'CustomDataSource',
          serviceKey: ServiceKey.create<IDataSource>('MyCompany:CustomDataSource', CustomDataSource)
      }
    ];
  }

  public name(): string {
    return 'MyCustomLibraryComponent';
  }
}
