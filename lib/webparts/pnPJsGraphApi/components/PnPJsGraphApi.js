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
import styles from './PnPJsGraphApi.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
var PnPJsGraphApi = /** @class */ (function (_super) {
    __extends(PnPJsGraphApi, _super);
    function PnPJsGraphApi() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    PnPJsGraphApi.prototype.render = function () {
        var _a = this.props, description = _a.description, isDarkTheme = _a.isDarkTheme, environmentMessage = _a.environmentMessage, hasTeamsContext = _a.hasTeamsContext, userDisplayName = _a.userDisplayName;
        return (React.createElement("section", { className: styles.pnPJsGraphApi + " " + (hasTeamsContext ? styles.teams : '') },
            React.createElement("div", { className: styles.welcome },
                React.createElement("img", { alt: "", src: isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png'), className: styles.welcomeImage }),
                React.createElement("h2", null,
                    "Well done, ",
                    escape(userDisplayName),
                    "!"),
                React.createElement("div", null, environmentMessage),
                React.createElement("div", null,
                    "Web part property value: ",
                    React.createElement("strong", null, escape(description)))),
            React.createElement("div", null,
                React.createElement("h3", null, "Welcome to SharePoint Framework!"),
                React.createElement("p", null, "The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling."),
                React.createElement("h4", null, "Learn more about SPFx development:"),
                React.createElement("ul", { className: styles.links },
                    React.createElement("li", null,
                        React.createElement("a", { href: "https://aka.ms/spfx", target: "_blank" }, "SharePoint Framework Overview")),
                    React.createElement("li", null,
                        React.createElement("a", { href: "https://aka.ms/spfx-yeoman-graph", target: "_blank" }, "Use Microsoft Graph in your solution")),
                    React.createElement("li", null,
                        React.createElement("a", { href: "https://aka.ms/spfx-yeoman-teams", target: "_blank" }, "Build for Microsoft Teams using SharePoint Framework")),
                    React.createElement("li", null,
                        React.createElement("a", { href: "https://aka.ms/spfx-yeoman-viva", target: "_blank" }, "Build for Microsoft Viva Connections using SharePoint Framework")),
                    React.createElement("li", null,
                        React.createElement("a", { href: "https://aka.ms/spfx-yeoman-store", target: "_blank" }, "Publish SharePoint Framework applications to the marketplace")),
                    React.createElement("li", null,
                        React.createElement("a", { href: "https://aka.ms/spfx-yeoman-api", target: "_blank" }, "SharePoint Framework API reference")),
                    React.createElement("li", null,
                        React.createElement("a", { href: "https://aka.ms/m365pnp", target: "_blank" }, "Microsoft 365 Developer Community"))))));
    };
    return PnPJsGraphApi;
}(React.Component));
export default PnPJsGraphApi;
//# sourceMappingURL=PnPJsGraphApi.js.map