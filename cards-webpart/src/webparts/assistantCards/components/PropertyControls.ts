import {
    PropertyPaneFieldType,
    IPropertyPaneField,
    IPropertyPaneCustomFieldProps
} from '@microsoft/sp-property-pane';

export class EmptyControl implements IPropertyPaneField<any> {
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public properties: IPropertyPaneCustomFieldProps;

    public render(elem: HTMLElement): void {
        elem.innerHTML = `<div style='margin-top:-4px'></div>`;
    }

    constructor(){
        this.properties = {
            onRender: this.render.bind(this),
            key: "Null"
        };
    }
}