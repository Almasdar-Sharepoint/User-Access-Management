import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane'; 
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as strings from 'UserAccessManagementWebPartStrings';
import UserAccessManagement from './components/UserAccessManagement';
import { IUserAccessManagementProps } from './components/IUserAccessManagement';
import { Constants } from './components/Constants/Constant';


// import * as msal from "@azure/msal-browser";

// let msalConfig : any; 
// let msalInstance : any;
// let tokenrequest : any;
// let currentAccount : any;

// let clientID: ''
// let authority: ''
// let scope: ''

// const msalConfig: msal.Configuration = {
//   auth: {
//     clientId: `${clientID}`,
//     authority: `${authority}`
//   },
//   cache: {
//     cacheLocation: "sessionStorage", // Configures cache location. "sessionStorage" is more secure, but "localStorage" gives you SSO between tabs.
//     storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge

//   },
//   system: {
//     iframeHashTimeout: 10000,
//     loggerOptions: {
//       loggerCallback: (level, message, containsPii) => {
//         if (containsPii) {
//           return;
//         }
//         switch (level) {
//           case msal.LogLevel.Error:
//             console.error(message);
//             return;
//           case msal.LogLevel.Info:
//             console.info(message);
//             return;
//           case msal.LogLevel.Verbose:
//             console.debug(message);
//             return;
//           case msal.LogLevel.Warning:
//             console.warn(message);
//             return;
//         }
//       },
//     },
//   },
// };


// const msalInstance: msal.PublicClientApplication = new msal.PublicClientApplication(
//   msalConfig
// );
// let currentAccount: msal.AccountInfo = null;
// const tokenrequest: any = {
//   scopes: `${scope}`,
//   //["Mail.Read"] ,//['api://c0f96cba-486b-4867-a2e2-e60ad996a3ed/hello'], ["Mail.Read"],
//   //scopes:["User.Read"],
//   account: currentAccount,
// };

export interface IUserAccessManagementWebPartProps {
  description: string;
  

}

export default class UserAccessManagementWebPart extends BaseClientSideWebPart<IUserAccessManagementWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private department: any = [];
  private AccessToken :any=""


  public render(): void {
    const element: React.ReactElement<IUserAccessManagementProps> = React.createElement(
      UserAccessManagement,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        department: this.department,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        email:this.context.pageContext.user.email,
        AccessToken: this.AccessToken
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    //await this.GetAuthentication();
    // await this.getAccessToken(this.context.pageContext.user.email);
    
    await this.GetELibrarywebAppAPI();

   
    const accountName = encodeURIComponent(`i:0#.f|membership|${this.context.pageContext.user.email}`);
    const userProfileUrl = `${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetUserProfilePropertyFor(accountName=@v,propertyName='Office')?@v='${accountName}'`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(userProfileUrl, SPHttpClient.configurations.v1);
    const userProfile = await response.json();
    const department = userProfile['value'];
    //console.log(department);
    this.department = department;
    //console.log(this.department);


    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  // public getAccessToken = async(email: string) => {
  //   try {
      
  //  sessionStorage.clear(); 
  //   localStorage.clear();
  //     // console.log(accounts);
  //     // this.GetIdTokenClaims();
  //     // let accessToken: string = null;
  //     await this.setCurrentAccount(email);//Sets the current account.
  //     // console.log(currentAccount);
  //     const loginRedirect: msal.RedirectRequest = tokenrequest;
  //     loginRedirect.loginHint = email;
  //     await msalInstance.acquireTokenPopup(loginRedirect)
  //       .then((tokenResponse : any) => {
  //         debugger;
  //         console.log(tokenResponse);
  //         localStorage.setItem("AccessToken", tokenResponse.accessToken);
          
  //         console.log(tokenResponse.accessToken);
  //         this.AccessToken=tokenResponse.accessToken
  //         //this.setState({ accessToken: tokenResponse.accessToken });
  //         if(tokenResponse.accessToken !== ""){
  //          // console.log(tokenResponse.accessToken, 'access token');
            
  //           Constants.AccessToken =  tokenResponse.accessToken;
  //         }
         

  //         // Constants.accessToken = tokenResponse.accessToken;
  //         // Constants.homeAccountId = tokenResponse.account.homeAccountId;
  //         // console.log("Token: ", Constants.accessToken);
  //         // return tokenResponse.accessToken;
  //       })
  //       // .then((tokenResponse)=>{
  //       //     console.log(tokenResponse);
  //       //     localStorage.setItem("AccessToken",tokenResponse)

  //       // })
  //       .catch((error: any) => {
  //         console.error(error);
  //         this.interactionRequired(email);


  //         // I haven't implemented redirect but it is fairly easy
  //         console.error("Maybe it is a popup blocked error. Implement Redirect");
  //         //return null;

  //       });

  //   } catch (error) {
  //     console.error(error);
  //   }
  // };

  // protected setCurrentAccount(email: string) {
  //   try {
  //     const currentAccounts: msal.AccountInfo[] = msalInstance.getAllAccounts();
  //     //returns all the accounts currently in the cache. An application must choose an account to acquire tokens silently.
  //     if (currentAccounts === null || currentAccounts.length === 0) {
  //       currentAccount = msalInstance.getAccountByUsername(
  //         email
  //       );
  //       //receives a username string and returns the matching account from the cache.
  //     } else if (currentAccounts.length > 1) {
  //       console.warn("Multiple accounts detected.");
  //       currentAccount = msalInstance.getAccountByUsername(
  //         email
  //       );
  //     } else if (currentAccounts.length === 1) {
  //       currentAccount = currentAccounts[0];
  //     }
  //     tokenrequest.account = currentAccount;
  //   } catch (error) {
  //     console.error(error);
  //   }
  // };
  // protected interactionRequired = (email: string): Promise<string> => {
  //   try {
  //     console.log("Inside Interaction");
  //     const loginPopupRequest: msal.PopupRequest = tokenrequest;
  //     loginPopupRequest.loginHint = email;
  //     return msalInstance
  //       //The resolution is to call an interactive method such as acquireTokenPopup
  //       .acquireTokenPopup(loginPopupRequest)
  //       .then((tokenResponse: any) => {
  //         localStorage.setItem("AccessToken", tokenResponse.accessToken);
  //         // Constants.accessToken = tokenResponse.accessToken;
  //         // Constants.homeAccountId = tokenResponse.account.homeAccountId;
  //         return tokenResponse.accessToken;
  //       })
  //       // .catch((error: any) => {
  //       //   console.error(error);
  //       //   // I haven't implemented redirect but it is fairly easy
  //       //   console.error("Maybe it is a popup blocked error. Implement Redirect");
  //       //   return null;
  //       // });
  //   } catch (error) {
  //     console.error(error);
  //   }
  // };

  


  public  GetELibrarywebAppAPI= async ()=> {
    try{
    let apiURL = `${this.context.pageContext.web.absoluteUrl}/_api/lists/GetByTitle('URL_Configuration')/items?$select=Title,WebAPiURL`;
    this.context.spHttpClient
      .get(apiURL, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          for (let i = 0; i < responseJSON.value.length; i++) {

          
              const check = responseJSON.value[i].Title
              //console.log(check, 'checkCondition')
              if (check === 'UserAccessManagementWebAPI') {
                Constants.var = responseJSON.value[i].WebAPiURL.Url;

              }
              else if(check === "UAMPowerBIUrl")
              {
                Constants.PowerBiReport = responseJSON.value[i].WebAPiURL.Url
                //console.log(Constants.PowerBiReport,"power bi report")
              }
            
          }
          //constants.shariaResearchWebapi =responseJSON.value[0].WebAPiURL.Url;
          //   this.setState({ allDataApi: responseJSON.value })
          //console.log(this.state.allDataApi, "web app apiurl");
        });
      });
    }
    catch(ex:any)
    {
      //console.log(ex);
    }
    // if(Webapi === "shariaResearchWebapi"){
    //   const ShariaResearchWebApi = this.state.allDataApi.map((obj:any)=>obj.WebAPiURL)
    // }
  }


  // public async GetAuthentication() {
  //   let apiURL = `${this.context.pageContext.web.absoluteUrl}/_api/lists/GetByTitle('ADAuthentication')/items?$select=Title,Value`;
  //   //alert(apiURL);
  //   await this.context.spHttpClient
  //     .get(apiURL, SPHttpClient.configurations.v1)
  //     .then((response: SPHttpClientResponse) => {
  //       response.json().then(async(responseJSON: any) => {
          


  //         let msalConfig: msal.Configuration = {
  //           auth: {
  //             clientId: `${responseJSON.value[0].Value}`,
  //             authority: `${responseJSON.value[1].Value}`
  //           },
  //           cache: {
  //             cacheLocation: "sessionStorage", // Configures cache location. "sessionStorage" is more secure, but "localStorage" gives you SSO between tabs.
  //             storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
          
  //           },
  //           system: {
  //             iframeHashTimeout: 10000,
  //             loggerOptions: {
  //               loggerCallback: (level, message, containsPii) => {
  //                 if (containsPii) {
  //                   return;
  //                 }
  //                 switch (level) {
  //                   case msal.LogLevel.Error:
  //                     console.error(message);
  //                     return;
  //                   case msal.LogLevel.Info:
  //                     console.info(message);
  //                     return;
  //                   case msal.LogLevel.Verbose:
  //                     console.debug(message);
  //                     return;
  //                   case msal.LogLevel.Warning:
  //                     console.warn(message);
  //                     return;
  //                 }
  //               },
  //             },
  //           },
  //         };
          
          
  //         let msalInstance1: msal.PublicClientApplication = new msal.PublicClientApplication(
  //           msalConfig
  //         );
  //         let currentAccount1: msal.AccountInfo = null;
  //         let tokenrequest1: any = {
  //           scopes: responseJSON.value[2].Value,
  //           //["Mail.Read"] ,//['api://c0f96cba-486b-4867-a2e2-e60ad996a3ed/hello'], ["Mail.Read"],
  //           //scopes:["User.Read"],
  //           account: currentAccount1,
  //         };
  //         debugger;
  //         console.log(tokenrequest1);
  //         console.log(currentAccount1);
          
          
          
  //         msalInstance = msalInstance1;
  //         currentAccount = currentAccount1;
  //         tokenrequest = tokenrequest1;
  //         // msalConfig = msalConfig1;

  //         await this.getAccessToken(this.context.pageContext.user.email);
          
  //             //  clientID = responseJSON.value[0].Value
  //             //  authority = responseJSON.value[1].Value
  //             //  scope = responseJSON.value[2].Value
  //            // console.log(check, 'checkCondition')
              
            
          
  //         //constants.shariaResearchWebapi =responseJSON.value[0].WebAPiURL.Url;
  //         //   this.setState({ allDataApi: responseJSON.value })
  //         //console.log(this.state.allDataApi, "web app apiurl");
  //       });
  //     });
  //   // if(Webapi === "shariaResearchWebapi"){
  //   //   const ShariaResearchWebApi = this.state.allDataApi.map((obj:any)=>obj.WebAPiURL)
  //   // }
  // }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
