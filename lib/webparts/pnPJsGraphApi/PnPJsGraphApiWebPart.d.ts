import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
export interface IPnPJsGraphApiWebPartProps {
    description: string;
}
export default class PnPJsGraphApiWebPart extends BaseClientSideWebPart<IPnPJsGraphApiWebPartProps> {
    private _isDarkTheme;
    private _environmentMessage;
    protected onInit(): Promise<void>;
    render(): void;
    private _renderEmailList;
    private _getEnvironmentMessage;
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=PnPJsGraphApiWebPart.d.ts.map