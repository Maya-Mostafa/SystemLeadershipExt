export class CustomComponentX extends React.Component<ICustomComponentProps, ICustomComponenState>  {

    constructor(props) {
        super(props);
        this.state = { 
            isPageRead: false,
            hideDialog: true
        };
    }

    public render() {

        // Parse custom object
        const myObject: IObjectParam = this.props.myObjectParam;
        
        console.log("pageContext", this.props.pageContext)

        return  (
            <div>
                <a onClick={() => this.setState({hideDialog: true})}>{this.props.pageTitleParam} --- </a>
                {this.state.isPageRead ?
                    <span>Yes, Page is Read!</span>
                    :
                    <span>Page is NOT Read Yet. Mark as <button onClick={() => updateMyUserProfile(this.context , ['111'], 'PDSBSystemLinks')}>Read!</button></span>
                }
                <IFrameDialog 
                    url={this.props.pageUrlParam}
                    hidden={this.state.hideDialog}
                    onDismiss={() => this.setState({hideDialog: true})}
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
            </div>
        )
    }
}