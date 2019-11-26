import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './UserAppsSampleWebPart.module.scss';
import * as strings from 'UserAppsSampleWebPartStrings';
import { IHttpClientOptions, HttpClientResponse, HttpClient, AadHttpClient, AadTokenProvider, AadTokenProviderFactory } from '@microsoft/sp-http';
import { getIconClassName } from '@uifabric/styling';

export interface IUserAppsSampleWebPartProps {
  description: string;
}

export default class UserAppsSampleWebPart extends BaseClientSideWebPart<IUserAppsSampleWebPartProps> {

  public render(): void {
    console.log("Starting render...");

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'webapps');

    callAPI(this);

    console.log("Ending render...");
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

function callAPI(me) {
  //establishes a timer that will throw an exception when it's likely that it has stalled out
  var timer = new Promise((resolve, reject) => { setTimeout(reject, 10000, 'API timed out.'); }); //10 seconds

  console.log("Getting token provider...");
  var getApps = me.context.aadTokenProviderFactory.getTokenProvider()
      .then(async (provider: AadTokenProvider) => { 
        console.log("Got token provider.");
        console.log(provider);

        console.log("Getting token...");
        return await provider.getToken('722dd073-dc67-40a6-b91e-969b294239be', false)
          .then(async (token) => { 
            console.log("Got token: " + token);
            return await GetApiData(me, token);
          }, (rejection) => {
            return Promise.reject("Token request rejected: " + rejection);
          });
      })
      .catch((err) => {
        return Promise.reject("Failed to get token provider: " + err);
      });
  

    // race() will get the response from whichever Promise resolves/rejects first.
    // In our case, the API call will run until the timer stops it with an automatic rejection.
    // This allows us to display an error message if it stalls out.
    Promise.race([getApps, timer]).then((value) => {
      console.log(value);        
    }).catch((err) => {
      console.log("Oops: " + err);
      me.context.statusRenderer.renderError(me.domElement, "Failed to retrieve list of apps.");
    });
}

async function GetApiData(me, token): Promise<string> {
  //makes the get() call that retrieves data from the API
  console.log("Calling API...");
  return await me.context.httpClient
    .get('https://www3.catholicmutual.org/SampleAPI/api/apps/', AadHttpClient.configurations.v1, {
      headers: [
        ['accept', 'application/json'],
        ['Authorization', 'Bearer ' + token]
      ]
    })
    .then((res: HttpClientResponse): Promise<any> => {
      console.log("Received a response from the API.");
      return res.json();
    })
    .then(data => {
      // process the data
      console.log(data);

      var listItems = ""; // stores the HTML links that will be displayed

      // create the HTML for each list item
      for(var i = 0; i < data.length; ++i){
        var app = data[i];
        
        var listItem = `<div role="listitem" class="${ styles.List_cell }">
          <div>
            <div>
              <a href="` + app.url + `" target="_blank" class="${ styles.ButtonCard }">
                <div class="${ styles.content }">
                  <div class="${ styles.thumbnail }">`;

                    if(app.icon != "") {
                      listItem += `<i class="${getIconClassName(app.icon)}" />`;
                    }
                    else {
                      listItem += `<i class="${getIconClassName('globe')}" />`;  // Generic Website
                    }                                    
                                              
                    listItem += `</i>
                    
                  </div>
                  <div class="${ styles.textArea }">
                    <div class="${ styles.labelTextWrapper }">
                      <div class="${ styles.lessText }">
                        ` + app.name + `
                      </div>
                    </div>
                  </div>
                </div>
              </a>
            </div>
          </div>
        </div>`;
        listItems += listItem;
      }

      // Display the final output
      me.domElement.innerHTML = `
        <div class="${ styles.userAppsSample }">
          <div class="${ styles.container }">
            <div class="${ styles.row }">
              ` + listItems + `
            </div>
          </div>
        </div>`;
        
      me.context.statusRenderer.clearLoadingIndicator(me.domElement);

      return "Successfully printed data from API."; // If we made it this far without an exception, it was a successful run.
    }, (err: any): void => {
      console.log("Error from API: " + err);
      me.context.statusRenderer.renderError(me.domElement, err);
    });
}
