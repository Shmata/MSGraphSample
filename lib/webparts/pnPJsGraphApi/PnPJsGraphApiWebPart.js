var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'PnPJsGraphApiWebPartStrings';
import PnPJsGraphApi from './components/PnPJsGraphApi';
var PnPJsGraphApiWebPart = /** @class */ (function (_super) {
    __extends(PnPJsGraphApiWebPart, _super);
    function PnPJsGraphApiWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._environmentMessage = '';
        return _this;
    }
    PnPJsGraphApiWebPart.prototype.onInit = function () {
        this._environmentMessage = this._getEnvironmentMessage();
        return _super.prototype.onInit.call(this);
    };
    PnPJsGraphApiWebPart.prototype.render = function () {
        var _this = this;
        this.context.msGraphClientFactory
            .getClient()
            .then(function (client) {
            // get information about the current user from the Microsoft Graph
            client
                //.api('/me/messages')
                .api('/users')
                .get(function (error, users, rawResponse) {
                console.log(error, users, rawResponse);
                _this.domElement.innerHTML = "\n            <div class=\"\">\n              <div>\n                  <h3>Welcome to SharePoint Framework!</h3>\n                  <p>\n                      The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.\n                  </p>\n              </div>\n              <div id=\"spListContainer\" />\n            </div>";
                // List the latest emails based on what we got from the Graph
                _this._renderEmailList(users.value);
            });
        });
        var element = React.createElement(PnPJsGraphApi, {
            description: this.properties.description,
            isDarkTheme: this._isDarkTheme,
            environmentMessage: this._environmentMessage,
            hasTeamsContext: !!this.context.sdks.microsoftTeams,
            userDisplayName: this.context.pageContext.user.displayName
        });
        ReactDom.render(element, this.domElement);
    };
    PnPJsGraphApiWebPart.prototype._renderEmailList = function (clinets) {
        console.log(clinets);
        var html = '';
        // for (let index = 0; index < clinets.length; index++) {
        //   html += `<p class="welcome">Email ${index + 1} - ${escape(clinets[index].mail)}</p>`;
        // }
        clinets.map(function (val, key) {
            html += "<p id='" + key + "' class=\"welcome\">Email: " + val.mail + " </p> <br />";
        });
        // Add the emails to the placeholder
        var listContainer = this.domElement.querySelector('#spListContainer');
        listContainer.innerHTML = html;
    };
    PnPJsGraphApiWebPart.prototype._getEnvironmentMessage = function () {
        if (!!this.context.sdks.microsoftTeams) { // running in Teams
            return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
        }
        return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
    };
    PnPJsGraphApiWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
        this.domElement.style.setProperty('--link', semanticColors.link);
        this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);
    };
    PnPJsGraphApiWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(PnPJsGraphApiWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    PnPJsGraphApiWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return PnPJsGraphApiWebPart;
}(BaseClientSideWebPart));
export default PnPJsGraphApiWebPart;
//# sourceMappingURL=PnPJsGraphApiWebPart.js.map