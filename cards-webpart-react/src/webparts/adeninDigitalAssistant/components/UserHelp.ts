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
                <div style="margin-bottom: 15px"><strong>New to Digital Assistant?</strong><br/>Head over to <a href="https://www.adenin.com/digital-assistant/" tabindex="-1">the adenin website</a> to learn more and create an account.</div>
                <div><strong>Need help setting up?</strong></br>Check out the <a href="https://www.adenin.com/support/">getting started guide</a> in the adenin developer documentation.</div>
            </div>`;
    }
}
export default UserHelp;