import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { DialogType, TooltipHost, IconButton, Icon } from 'office-ui-fabric-react';
import {MSGraphClientFactory, SPHttpClient} from "@microsoft/sp-http";
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { updateMyUserProfile, getmyUserProfileProps, getMyPropIds, getDefaultTaskListID, addToTasks } from './Services/DataRequests';
import { NewTask } from './NewTask/NewTask';
import spservices from './Services/spservices';
import SendEmail from './SendEmail/SendEmail';
import { sp} from '@pnp/pnpjs';
import {Web} from "@pnp/sp/presets/all";
import { spfi, SPFx} from "@pnp/sp";
import "@pnp/sp/webs";
import { AssignFrom } from "@pnp/core";
import { Caching } from "@pnp/queryable";

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
    msGraphClientFactory?: MSGraphClientFactory;
    pages?: any;
}

export function CustomComponent (props: ICustomComponentProps){


    //const sp = spfi().using();
    //const web = Web([sp.web, 'https://pdsb1.sharepoint.com']);

    const web = Web("https://pdsb1.sharepoint.com").using(SPFx({pageContext: props.pageContext})).using(Caching());
    console.log("web", web);

    const dateOptions: any = { year: 'numeric', month: 'long', day: 'numeric' };

    const profilePropSystemLinks = 'PDSBSystemLinks';
    const [userSysLinksPropIds, setUserSysLinksPropIds] = React.useState(new Set());

    const profilePropTodos = 'PDSBMyTodos';
    const [userTodosIds, setUserTodosIds] = React.useState(new Set());

    const [hideDialog, setHideDialog] = React.useState(true);
    const [pageUrl, setPageUrl] = React.useState('');

    const [showPlannerDlg, setShowPlannerDlg] = React.useState(false);
    const [taskDetails, setTaskDetails] = React.useState(null);
    const _spservices = new spservices(props.pageContext, props.msGraphClientFactory);
    //const _spservices = null;

    const [showEmailDlg, setShowEmailDlg] = React.useState(false);

    React.useEffect(()=>{
        getmyUserProfileProps(props.sphttpClient).then(myUserProfileProps => {
            const myPropsSysLinksIds = getMyPropIds(myUserProfileProps, profilePropSystemLinks);
            setUserSysLinksPropIds(myPropsSysLinksIds);

            const myPropsTodosIds = getMyPropIds(myUserProfileProps, profilePropTodos);
            setUserTodosIds(myPropsTodosIds);
        });        
    }, []);

    React.useEffect(()=>{
        // console.log("userEffect run!");
        // console.log("userPropsIds string", Array.from(userPropIds).toString());
    }, [Array.from(userSysLinksPropIds).toString(), Array.from(userTodosIds).toString()]);

    const checkHandler = (pageId: string) => {
        const cloneIds = new Set(userSysLinksPropIds);
        cloneIds.add(pageId);
        setUserSysLinksPropIds(cloneIds);
        updateMyUserProfile(props.pageContext, props.sphttpClient, cloneIds, profilePropSystemLinks);
    };
    const unCheckHandler = (pageId: string) => {
        const cloneIds = new Set(userSysLinksPropIds);
        cloneIds.delete(pageId);
        setUserSysLinksPropIds(cloneIds);
        updateMyUserProfile(props.pageContext, props.sphttpClient, cloneIds, profilePropSystemLinks);
    };

    const dialogOpenHandler = (link: string) => {
        setPageUrl(link);
        setHideDialog(false);
    };

    const addTodoHandler = (page: any) => {
        console.log("addTodoHandler");
        if (!userTodosIds.has(page.ListItemID)){
            console.log("addTodoHandler active");
            getDefaultTaskListID(props.msGraphClientFactory).then(resId => {
                addToTasks(props.msGraphClientFactory, resId, page);
                
                const cloneIds = new Set(userTodosIds);
                cloneIds.add(page.ListItemID);
                setUserTodosIds(cloneIds);
                updateMyUserProfile(props.pageContext, props.sphttpClient, cloneIds, profilePropTodos);
            });
        }
    };

    const taskDetailsPlannerHandler = (pageDetails: any) => {
        setTaskDetails(pageDetails);
        setShowPlannerDlg(true);
    };

    const sendEmailHandler = (pageDetails: any) => {
        setTaskDetails(pageDetails);
        setShowEmailDlg(true);
    };

    const spTestFncs = () => {
        _spservices.getUserGroups().then(res => console.log("---- spServices : geUserGroups ----", res));
        _spservices.getUserPlansByGroupId('acbcf16c-c862-4c61-ae32-8f629366451a').then(res => console.log("---- spServices : getUserPlansByGroupId ----", res)); //Portal & Collaboration
    };



    // console.log(props.pages);
    // console.log(props.pageContext);

    return (
        <>
            {/* <button onClick={spTestFncs}>Test sp functions</button> */}
            <ul className='template--defaultList'>
                {props.pages.items.map(page => {
                    return (
                        <li className='template--listItem'>
                            {!userSysLinksPropIds.has(page.ListItemID) 
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
                                        <a data-interception="off" target='_blank' href={page.Path} className='page-link'>{page.Title}</a>
                                        {/* <a className='page-link' onClick={() => dialogOpenHandler(page.Path)}>{page.Title}</a> */}
                                        {/* <a data-interception="off" className='page-link-new-window' target='_blank' href={page.Path}><Icon iconName='OpenInNewWindow' /></a> */}
                                    </span>
                                    <span>
                                        {page.AuthorOWSUSER &&<span className='template--listItem--author'>{page.AuthorOWSUSER.split('|')[1]}</span>}
                                        <span className='template--listItem--date'>{new Date(page.Created).toLocaleDateString('en-us', dateOptions)}</span>
                                    </span>
                                    {(page.TaskDueDateOWSDATE || page.RefinableString110) && <span className='due-date'><Icon iconName='Calendar' />Due by: {page.TaskDueDateOWSDATE || page.RefinableString110}</span> }
                                    {page.RefinableString137 &&  <div>Attachments: {page.RefinableString137}</div>}
                                
                                    <div className='actions'>
                                        <button onClick={() => sendEmailHandler(page)}><img width='20' src={require('./icons/Outlook.svg')} />Send by E-mail</button>
                                        <button className={!userTodosIds.has(page.ListItemID) ? '' : 'actionDisabled'} onClick={()=> addTodoHandler(page)}><img width='20' src={require('./icons/Todo.svg')} />{!userTodosIds.has(page.ListItemID) ? 'Add' : 'Added'} to Todo</button>
                                        <button onClick={() => taskDetailsPlannerHandler(page)}><img width='20' src={require('./icons/Planner.svg')} />Add to Planner</button>
                                    </div>
                                </div>
                            </div>
                            <div className='template--listItem--thumbnailContainer'>
                                <div className='thumbnail--image'>
                                    <img width="120" src={page.AutoPreviewImageUrl} />
                                </div>
                            </div>
                        </li>
                    );
                })}
            </ul>
            {showPlannerDlg &&
                <NewTask 
                    displayDialog = {showPlannerDlg} 
                    onDismiss = {() => setShowPlannerDlg(false)} 
                    spservice = {_spservices} 
                    taskDetails = {taskDetails}
                />
            }
            {showEmailDlg &&
                <SendEmail 
                    displayDialog = {showEmailDlg} 
                    onDismiss = {() => setShowEmailDlg(false)} 
                    taskDetails = {taskDetails}
                    context = {web}
                />
            }
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
    private pageContext: PageContext;
    private msGraphClientFactory: MSGraphClientFactory;

    public constructor() {
        super(); 
        this._serviceScope.whenFinished(()=>{
            this.pageContext = this._serviceScope.consume(PageContext.serviceKey);
            this.sphttpClient = this._serviceScope.consume(SPHttpClient.serviceKey);
            this.msGraphClientFactory = this._serviceScope.consume(MSGraphClientFactory.serviceKey);
        });
    }
 
    public async connectedCallback() {
        let props = this.resolveAttributes();
        const customComponent = <CustomComponent pageContext={this.pageContext} sphttpClient={this.sphttpClient} msGraphClientFactory={this.msGraphClientFactory} {...props}/>;
        ReactDOM.render(customComponent, this);
    }    
}

