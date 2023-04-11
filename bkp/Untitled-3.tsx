import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { DialogType } from 'office-ui-fabric-react';
import * as ReactDOM from 'react-dom';
import {SPHttpClient, ISPHttpClientOptions} from "@microsoft/sp-http";
import { PageContext } from '@microsoft/sp-page-context';

export interface IObjectParam {
    myProperty: string;
}

export interface ICustomComponentProps {

    /**
     * A sample string param
     */
    myStringParam?: string;

    /**
     * A sample object param
     */
    myObjectParam?: IObjectParam;

    /**
     * A sample date param
     */
    myDateParam?: Date;

    /**
     * A sample number param
     */
    myNumberParam?: number;

    /**
     * A sample boolean param
     */
    myBooleanParam?: boolean;

    pageUrlParam? : string;
    pageTitleParam? : string;
    pageFileTypeParam? : string;   
    pageContext?: any; 
}


export interface ICustomComponenState {
}

export const updateMyUserProfile = async (context: any, listItems: any, profilePropName: string) =>{
    //const updatedIds = getUpdatedProfileIds(listItems);
    const updatedIds = listItems;

    console.log(context)
    
    const responseUrl = `https://pdsb1.sharepoint.com/_api/SP.UserProfiles.PeopleManager/SetMultiValuedProfileProperty` ;

    let userData = {
        'accountName': "i:0#.f|membership|" + context.pageContext.user.email,
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

    const _data = await context.spHttpClient.post(responseUrl, SPHttpClient.configurations.v1, spOptions);
    if (_data.ok){
        console.log('User Profile property '+profilePropName+' is updated!');
    }
};

export function CustomComponent (props: ICustomComponentProps){

    const [hideDialog, setHideDialog] = React.useState(true);
    const [isPageRead, setIsPageRead] = React.useState(false);

    console.log("props.pageContext", props.pageContext);

    return (
        <div>
            <a onClick={() => setHideDialog(false)}>{props.pageTitleParam} --- </a>
            {isPageRead ?
                <span>Yes, Page is Read!</span>
                :
                <span>Page is NOT Read Yet. Mark as <button onClick={() => updateMyUserProfile(props.pageContext, ['111'], 'PDSBSystemLinks')}>Read!</button></span>
            }
            <IFrameDialog 
                url={props.pageUrlParam}
                hidden={hideDialog}
                onDismiss={() => setHideDialog(true)}
                modalProps={{
                    isBlocking: true,
                }}
                dialogContentProps={{
                    type: DialogType.close,
                    showCloseButton: true
                }}
                width={'800px'}
                height={'600px'}/>
        </div>
    )
}

export class MyCustomComponentWebComponent extends BaseWebComponent {
    
    public constructor() {
        super(); 
    }
 
    public async connectedCallback() {
        
        let props = this.resolveAttributes();
        const customComponent = <CustomComponent {...props}/>;
        ReactDOM.render(customComponent, this);
    }    
}