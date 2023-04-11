import {SPHttpClient, ISPHttpClientOptions} from "@microsoft/sp-http";

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

export const updateMyUserProfile = async (context: any, sphttpClient: any,  listItems: any, itemID: string,  profilePropName: string) =>{
    
    console.log("listItems", listItems);
    const updatedIds = Array.from(listItems.add(itemID));
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