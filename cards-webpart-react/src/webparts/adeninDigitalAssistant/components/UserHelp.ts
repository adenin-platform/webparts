import { IPropertyPaneField, PropertyPaneFieldType, IPropertyPaneCustomFieldProps } from "@microsoft/sp-property-pane";

export class UserHelp implements IPropertyPaneField<IPropertyPaneCustomFieldProps> {
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public properties: IPropertyPaneCustomFieldProps;

    constructor() {
         this.properties = {
             key: "Logo",
             onRender: this.onRender.bind(this)
        };
    }

    private onRender(elem: HTMLElement): void {
        elem.innerHTML = `
            <div style="margin: 30px 0;">
                <div style="margin-bottom: 15px"><strong>New to Digital Assistant?</strong><br/><a href="https://www.adenin.com/digital-assistant/" tabindex="-1">Digital Assistant</a> is a free platform that connects to all the business apps you already use.<br/><a href="https://www.adenin.com/digital-assistant/" tabindex="-1">Click here to create your Assistant</a></div>
                <div><strong>Looking for setup instructions?</strong></br>Check out <a href="https://www.adenin.com/support/topic/?topic=SharePoint">resources from adenin Support</a> to help you get started.</div>
            </div>`;
    }
}
export default UserHelp;