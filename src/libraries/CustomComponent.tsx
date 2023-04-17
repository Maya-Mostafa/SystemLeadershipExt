import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { DialogType, TooltipHost, IconButton, Icon } from 'office-ui-fabric-react';
import * as ReactDOM from 'react-dom';
import {SPHttpClient} from "@microsoft/sp-http";
import { PageContext } from '@microsoft/sp-page-context';
import { updateMyUserProfile, getmyUserProfileProps, getMyPropIds } from './Services/DataRequests';

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

    const dateOptions = { year: 'numeric', month: 'long', day: 'numeric' };

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

    const checkHandler = (pageId: string) => {
        const cloneIds = new Set(userPropIds);
        cloneIds.add(pageId);
        setUserPropIds(cloneIds);
        updateMyUserProfile(props.pageContext, props.sphttpClient, cloneIds, profilePropName);
    };
    const unCheckHandler = (pageId: string) => {
        const cloneIds = new Set(userPropIds);
        cloneIds.delete(pageId);
        setUserPropIds(cloneIds);
        updateMyUserProfile(props.pageContext, props.sphttpClient, cloneIds, profilePropName);
    };

    const dialogOpenHandler = (link: string) => {
        setPageUrl(link);
        setHideDialog(false);
    };

    console.log(props.pages);
    console.log(props.pageContext);

    return (
        <>
            <ul className='template--defaultList'>
                {props.pages.items.map(page => {
                    return (
                        <li className='template--listItem'>
                            {!userPropIds.has(page.ListItemID) 
                                ?
                                <TooltipHost content="Check done" calloutProps={{ gapSpace: 0 }}>
                                    <IconButton className='uncheck-btn' onClick={() => checkHandler(page.ListItemID)} iconProps={{ iconName: 'Accept' }} />
                                </TooltipHost>
                                :
                                <TooltipHost content="Uncheck done" calloutProps={{ gapSpace: 0 }}>
                                    <IconButton className='check-btn' onClick={() => unCheckHandler(page.ListItemID)} iconProps={{ iconName: 'Accept' }} />
                                </TooltipHost>
                            }
                            <div className='template--listItem--result'>
                                <div className='template--listItem--contentContainer'>
                                    {page.RefinableString129 && <span className="dept-hdr">{page.RefinableString129}</span>} 
                                    <span className='template--listItem--title example-themePrimary'>
                                        <a className='page-link' onClick={() => dialogOpenHandler(page.Path)}>{page.Title}</a>
                                        <a data-interception="off" className='page-link-new-window' target='_blank' href={page.Path}><Icon iconName='OpenInNewWindow' /></a>
                                    </span>
                                    <span>
                                        <span className='template--listItem--author'>{page.AuthorOWSUSER.split('|')[1]}</span>
                                        <span className='template--listItem--date'>{new Date(page.Created).toLocaleDateString('en-us', dateOptions)}</span>
                                    </span>
                                    {(page.TaskDueDateOWSDATE || page.RefinableString110) && <span className='due-date'><Icon iconName='Calendar' />Due by: {page.TaskDueDateOWSDATE || page.RefinableString110}</span> }
                                    {page.RefinableString137 &&  <div>Attachments: {page.RefinableString137}</div>}
                                </div>
                            </div>
                            <div className='template--listItem--thumbnailContainer'>
                                <div className='thumbnail--image'>
                                    <img width="120" src={page.AutoPreviewImageUrl} />
                                </div>
                            </div>
                        </li>
                    )
                })}
            </ul>
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