import * as React from 'react';
import ISendEmailProps  from './ISendEmailProps';
import { Dialog, DialogType, DialogFooter, PrimaryButton, DefaultButton, ContextualMenu, Label, TextField } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import styles from './SendEmail.module.scss';
import {sendEmailGraph} from '../Services/DataRequests';

export default function SendEmail (props: ISendEmailProps){

    console.log("SendEmail props", props);

    const dragOptions = {
        moveMenuItemText: 'Move',
        closeMenuItemText: 'Close',
        menu: ContextualMenu,
    };
    const dialogContentProps = {
        type: DialogType.normal,
        title: 'Send this Page by email',
    };

    const [opMsg, setOpMsg] = React.useState('');
    const [recipients, setRecipients] = React.useState([]);

    const getPeoplePickerItems = (items: any[]) => {
        // console.log("people picker items", items);
        setRecipients(items.map(item => item.secondaryText));
    };

    const messageBody = 
        <div className={styles.sendEmail}>
            <a href={props.taskDetails.Path}>
                <img height={135} src={props.taskDetails.AutoPreviewImageUrl} alt={props.taskDetails.Title} />
            </a>
            <a href={props.taskDetails.Path}>
                {props.taskDetails.Title}
            </a>
            <p>{props.taskDetails.HitHighlightedSummary}</p>
        </div>
    ;

    const emailBody = `
        <div style="width: 100%;">
            <table align="center" cellspacing="0" cellpadding="0" border="0" with="600">
                <tr>
                    <td width="540" colspan="2">
                        <span style="display: inline-block;color:#333333; font-family:'Segoe UI','Segoe UI',Tahoma,Arial,sans-serif; font-size:15px; font-weight:normal; padding:20px">${opMsg}</span>
                    </td>
                </tr>
                <tr>                
                    <td width="192" style="width:192px; vertical-align:top; padding-left:20px">
                        <a style="display:inline-block" href='${props.taskDetails.Path}'>
                            <div>
                                <img alt="${props.taskDetails.Title}" title="${props.taskDetails.Title}" width="192" height="128" style="min-width: auto; min-height: auto;display:block" src='${props.taskDetails.PictureThumbnailURL}'/>
                            </div>
                        </a>
                    </td>
                    <td width="348" style="width:348px; vertical-align:top; padding:0 20px">
                        <a style="margin-bottom:20px;display: inline-block;color:#0078d7;font-weight:400;font-size:17px;font-family: 'Segoe UI',Tahoma,Arial,sans-serif;text-decoration: none;word-break: break-word;" href='${props.taskDetails.Path}'>
                            ${props.taskDetails.Title}
                            <br/>
                        </a>
                        <div style="line-height: 18px;color: #666666;font-weight: 400;font-size: 13px;font-family: 'Segoe UI',Tahoma,Arial,sans-serif;word-break: break-word;">${props.taskDetails.HitHighlightedSummary}</div>
                    </td>
                </tr>
            </table>
            <div style="height: 103px; width: 100%;">
                <table width="600" cellspacing="0" cellpadding="0" border="0" style="border-collapse: separate; padding: 45px 20px 0px; transform: scale(0.915, 0.915); transform-origin: left top;">
                    <tbody>
                        <tr>
                            <td width="600" align="left" valign="top">
                                <table border-collapse:separate; border-top:0.5px solid #b9b9b9; padding:12px 0 24px 0>
                                    <tbody>
                                        <tr>
                                        </tr>
                                    </tbody>
                                </table>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    `;
    
    const sendEmailHandler = () => {
        if (recipients.length !== 0){
            sendEmailGraph(props.context.msGraphClientFactory, recipients, props.taskDetails.Title, emailBody);
            props.onDismiss();
        }
    };

    const peoplePickerErrorHandler = (items: any[]) => {
        if (items.length === 0)
            return 'This field is required'
    };

    return(
        <>
            <Dialog
                hidden={!props.displayDialog}
                onDismiss={props.onDismiss}
                dialogContentProps={dialogContentProps}
                minWidth={550}
            >
                <form>
                    <Label>To : </Label>
                    <PeoplePicker
                        context={props.context}
                        personSelectionLimit={10}
                        groupName={''}
                        resolveDelay={1000}
                        ensureUser={true}
                        onChange={getPeoplePickerItems}
                        required
                        // errorMessage={'This field is required'}
                        validateOnFocusOut
                        onGetErrorMessage={peoplePickerErrorHandler}
                    /> 
                    <br />
                    <TextField 
                        multiline 
                        rows={5} 
                        placeholder='Add a message (optional)' 
                        value={opMsg}
                        onChange={(e:any) => setOpMsg(e.target.value)}
                    />
                    {messageBody}
                </form>
                <DialogFooter>
                    <PrimaryButton onClick={sendEmailHandler} text="Send" />
                    <DefaultButton onClick={props.onDismiss} text="Cancel" />
                </DialogFooter>
            </Dialog>
        </>
    );
}
