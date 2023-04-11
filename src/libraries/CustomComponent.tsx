import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { DialogType } from 'office-ui-fabric-react';
import * as ReactDOM from 'react-dom';
import {SPHttpClient} from "@microsoft/sp-http";
import { PageContext } from '@microsoft/sp-page-context';
import { updateMyUserProfile, getmyUserProfileProps, getMyPropIds } from './Services/DataRequests';

import ReadStatus from './ReadStatus';

export interface IObjectParam {
    myProperty: string;
}

export interface ICustomComponentProps {
    pageUrlParam? : string;
    pageTitleParam? : string;
    pageFileTypeParam? : string;   
    pageId?: string;
    pageContext?: PageContext; 
    sphttpClient?: SPHttpClient;
    pages?: any;
}

export function CustomComponent (props: ICustomComponentProps){

    const profilePropName = 'PDSBSystemLinks';
    const [hideDialog, setHideDialog] = React.useState(true);
    const [userPropIds, setUserPropIds] = React.useState(new Set());
    const [pageUrl, setPageUrl] = React.useState('');

    React.useEffect(()=>{
        getmyUserProfileProps(props.sphttpClient).then(myUserProfileProps => {
            const myPropsIds = getMyPropIds(myUserProfileProps, profilePropName);
            setUserPropIds(myPropsIds);
        });
    }, []);

    React.useEffect(()=>{
        // console.log("userEffect run!");
        // console.log("userPropsIds string", Array.from(userPropIds).toString());
    }, [Array.from(userPropIds).toString()]);

    const readHandler = (pageId: string) => {
        setUserPropIds(prev => {
            const cloneIds = new Set(prev);
            return cloneIds.add(pageId);
        });
        updateMyUserProfile(props.pageContext, props.sphttpClient, userPropIds, pageId, profilePropName);
    };

    const dialogOpenHandler = (link: string) => {
        setPageUrl(link);
        setHideDialog(false);
    };

    console.log(props.pages);

    return (
        <>
            <table cellPadding={5} style={{textAlign: 'left'}}> 
                <tr>
                    <th>ID</th>
                    <th>Title</th>
                    <th>Open</th>
                    <th>is Read</th>
                    <th>Action</th>
                </tr>
                {props.pages.items.map(page => {
                    return (
                        <tr>
                            <td>{page.ListItemID}</td>
                            <td>
                                <a target='_blank' href={page.Path}>{page.Title}</a>
                            </td>
                            <td>
                                <button onClick={() => dialogOpenHandler(page.Path)}>Open In Dialog</button>
                            </td>
                            <td>
                                <ReadStatus listItemID={page.ListItemID} userPropIds={userPropIds} />
                            </td>
                            <td>
                                {!userPropIds.has(page.ListItemID) && 
                                    <button onClick={() => readHandler(page.ListItemID)}>Read!</button>
                                }
                            </td>
                        </tr>
                    );
                })}
            </table>
            <IFrameDialog 
                url={pageUrl}
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
                height={'600px'}
            />
        </>
    );
}

export class MyCustomComponentWebComponent extends BaseWebComponent {
    
    private sphttpClient: SPHttpClient;
    private pageContext:PageContext;

    public constructor() {
        super(); 
        this._serviceScope.whenFinished(()=>{
            this.pageContext = this._serviceScope.consume(PageContext.serviceKey);
            this.sphttpClient= this._serviceScope.consume(SPHttpClient.serviceKey);
            //const msGraphClientFactory: MSGraphClientFactory= this._serviceScope.consume(MSGraphClientFactory.serviceKey)
        });
    }
 
    public async connectedCallback() {
        let props = this.resolveAttributes();
        const customComponent = <CustomComponent pageContext={this.pageContext} sphttpClient={this.sphttpClient} {...props}/>;
        ReactDOM.render(customComponent, this);
    }    
}