import { BaseComponentContext } from "@microsoft/sp-component-base";

export default interface ISendEmailProps{
    displayDialog:boolean;
    onDismiss: () => void;
    taskDetails: any;
    context: BaseComponentContext;
}