import {SPHttpClient, ISPHttpClientOptions} from "@microsoft/sp-http";

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/sputilities";
import { IEmailProperties } from "@pnp/sp/sputilities";

export const sendEmailPnP = async (context: any, recipient: string, sender: string, taskDetails: any, body: string) => {
    const sp = spfi().using(SPFx(context));
    const emailProps: IEmailProperties = {
        To: recipient.split(';'),
        From: sender,
        Subject: taskDetails.Title,
        Body: body,
        AdditionalHeaders: {
            "content-type": "text/html"
        }
    };

    await sp.utility.sendEmail(emailProps);
    console.log("Email Sent!");

};

export const sendEmailSP = async (context: any, sphttpClient: any) => {

    const siteUrl = context.pageContext._web.serverRelativeUrl;
    const emailUrl = siteUrl + '/_api/SP.Utilities.Utility.SendEmail';

    console.log("emailUrl", emailUrl);

    const   reqOptions: ISPHttpClientOptions  = {
            headers: {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "odata-version":"3.0",
            },
            body: JSON.stringify({
                'properties': {
                    '__metadata': { 'type': 'SP.Utilities.EmailProperties' },
                    'To': { 'results': ['mai.mostafa@peelsb.com'] },
                    'Body': 'Ignore this test email.',
                    'Subject': 'Subject Test Email',
                    'From' : 'mai.mostafa@peelsb.com'
                }
            })
    };

    const _data = await sphttpClient.post(emailUrl, SPHttpClient.configurations.v1, reqOptions);
    if (_data.ok){
        console.log('Email Sent!');
    }
};

export const sendEmailGraph = async (msGraphClientFactory: any, recipients: any, subject: string, body: string) => {
    
    const emailARs = recipients.map(item => { return {emailAddress: {address: item}} });

    const sendMail = {
        message: {
          subject: subject,
          body: {
            contentType: 'HTML',
            content: body
          },
          toRecipients: emailARs,
        },
        saveToSentItems: 'true'
      };

      const graphClient = await msGraphClientFactory.getClient();
      const graphPostResponse = await graphClient.api("/me/sendMail").post(sendMail);

      return graphPostResponse;
};


export const getmyUserProfileProps = async (sphttpClient: any) => {
    const responseUrl = `https://pdsb1.sharepoint.com/_api/SP.UserProfiles.PeopleManager/GetMyProperties` ;
    
    try{
        const response = await sphttpClient.get(responseUrl, SPHttpClient.configurations.v1);
        if (response.ok){
            const responseResults = await response.json();
            return responseResults.UserProfileProperties;
        }else{
            console.log("User Profile props Error: " + response.statusText);
        }
    }catch(error){
        console.log("User Profile props Response Error: " + error);
    }
};

export const getMyPropIds = (myUserProfileProps: any, profilePropName: string) => {
    for (let userProp of myUserProfileProps){
        if (userProp.Key === profilePropName){
            return new Set(userProp.Value.split('|'));
        }
    }
};

export const updateMyUserProfile = async (context: any, sphttpClient: any,  listItems: any, profilePropName: string) =>{
    
    console.log("listItems", listItems);
    const updatedIds = Array.from(listItems);
    console.log("updatedIds", updatedIds);

    const responseUrl = `https://pdsb1.sharepoint.com/_api/SP.UserProfiles.PeopleManager/SetMultiValuedProfileProperty` ;

    let userData = {
        'accountName': "i:0#.f|membership|" + context._user.email,
        'propertyName': profilePropName,
        'propertyValues': updatedIds
    },
    spOptions: ISPHttpClientOptions = {
        headers:{
            "Accept": "application/json;odata=nometadata", 
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
        },
        body: JSON.stringify(userData)
    };

    const _data = await sphttpClient.post(responseUrl, SPHttpClient.configurations.v1, spOptions);
    if (_data.ok){
        console.log('User Profile property '+profilePropName+' is updated!');
    }
};

export const getDefaultTaskListID = async (msGraphClientFactory: any) => {
    const grapClient = await msGraphClientFactory.getClient();
    const graphGetResponse = await grapClient.api("/me/todo/lists").get();
    return graphGetResponse.value.filter(item => item.wellknownListName === 'defaultList')[0].id;
};

export const addToTasks = async (msGraphClientFactory: any, defaultTaskListId: string, page: any) => {

    const todoTask = {
        title: page.Title,
        body: {
            "content": `${page.Title} (${page.RefinableString129}) - ${page.Path}`,
            "contentType": "text"
        },
        linkedResources: [
           {
              webUrl: page.Path,
              applicationName: 'System Leadership News',
            //   displayName: 'For Action: PDSB’s Response to the Right To Read: Shifting Literacy Instruction and Assessment in Peel Updated'
           }
        ]
     };

    const grapClient = await msGraphClientFactory.getClient();
    const graphGetResponse = await grapClient.api(`/me/todo/lists/${defaultTaskListId}/tasks`).post(todoTask);
    console.log("Todo Task Created!");
    return graphGetResponse;
};