import * as React from 'react';
import ISendEmailProps  from './ISendEmailProps';
import { Dialog, DialogType, DialogFooter, PrimaryButton, DefaultButton, ContextualMenu, Label, TextField } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export default function SendEmail (props: ISendEmailProps){

    console.log("SendEmail props", props);

    const dragOptions = {
        moveMenuItemText: 'Move',
        closeMenuItemText: 'Close',
        menu: ContextualMenu,
    };
    const modalPropsStyles = { 
        main: { maxWidth: 450 } 
    };
    const dialogContentProps = {
        type: DialogType.normal,
        title: 'Send this page by email',
    };

    return(
        <>
            <Dialog
                hidden={!props.displayDialog}
                onDismiss={props.onDismiss}
                dialogContentProps={dialogContentProps}
            >
                <form>
                    <Label>To : </Label>
                    <PeoplePicker
                        context={props.context}
                        personSelectionLimit={10}
                        showHiddenInUI={false}
                        groupName={''}
                        resolveDelay={1000}
                        ensureUser={true}
                    />
                    <br />
                    <Label>CC : </Label>
                    {/* <PeoplePicker
                        context={props.context}
                        personSelectionLimit={10}
                        showHiddenInUI={false}
                        groupName={''}
                        resolveDelay={1000}
                        ensureUser={true}
                    /> */}
                    <br />
                    <Label>Subject : </Label>
                    <TextField
                        // value={this.state.EmailSubject}
                        // onChange={this._onChange1}
                    />
                    <br />
                    <Label>Body : </Label>
                    <TextField multiline rows={3} placeholder="Body.." />
                </form>
                <DialogFooter>
                    <PrimaryButton onClick={props.onDismiss} text="Send" />
                    <DefaultButton onClick={props.onDismiss} text="Cancel" />
                </DialogFooter>
            </Dialog>
        </>
    );
}
