export class CustomComponent extends React.Component<ICustomComponentProps, ICustomComponenState> {

    
    
    public render() {

        // Parse custom object
        const myObject: IObjectParam = this.props.myObjectParam;

        return  <div>
                
                    <div>{this.props.pageTitleParam}</div>

                    {/* {this.props.myStringParam} {myObject.myProperty} */}
                    <IFrameDialog 
                        url={this.props.pageUrlParam}
                        // iframeOnLoad={this._onIframeLoaded.bind(this)}
                        hidden={true}
                        // onDismiss={this._onDialogDismiss.bind(this)}
                        modalProps={{
                            isBlocking: true,
                        }}
                        dialogContentProps={{
                            type: DialogType.close,
                            showCloseButton: true
                        }}
                        width={'570px'}
                        height={'315px'}/>
                </div>;
    }
}