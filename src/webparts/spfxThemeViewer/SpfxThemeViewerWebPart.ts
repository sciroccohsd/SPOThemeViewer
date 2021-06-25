import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneLabel,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxThemeViewerWebPart.module.scss';
import * as strings from 'SpfxThemeViewerWebPartStrings';

export interface ISpfxThemeViewerWebPartProps {
    description: string;
}

declare global {
    interface Window {
        __themeState__: any;
    }
}
    
export default class SpfxThemeViewerWebPart extends BaseClientSideWebPart<ISpfxThemeViewerWebPartProps> {


    public render(): void {
        const root = this.CreateElement("div", { "class": styles.spfxThemeViewer }),
            warning = this.CreateElement("div", { "class": styles.warning, "text": "NOTE: Not all theme names are available to every SPO theme set." }),
            msg = this.CreateElement("div", { "text": "Theme values from: window.__themeState__.theme" }),
            link = this.CreateElement("a", {
                "class": styles.link,
                "href": "https://docs.microsoft.com/en-us/sharepoint/dev/design/design-guidance-overview",
                "target": "_blank",
                "text": "Designing great SharePoint experiences - Overview (MS Docs)"
            }),
            moreInfo = this.CreateElement("div", { "text": 'Example: background-color: "[theme: bodyBackground, default:#ffffff]";' }),
            names = Object.keys(window.__themeState__.theme);

        names.sort().forEach(name => {
            const value = window.__themeState__.theme[name];
            root.append(this.CreateThemeBox(name, value));
        });

        this.domElement.append(warning, msg, link, moreInfo, root);
    }

    private CreateThemeBox(name: string, value: string): HTMLDivElement {
        const wrapper = this.CreateElement("div", { "class": styles.boxWrapper }),
            label = this.CreateElement("div", { "text": name }),
            colorbox = this.CreateElement("div", {
                "style": `background-color: ${ value }`,
                "class": styles.box
            }),
            display = this.CreateElement("div", {
                "class": `themeviewer-display ${ styles.display }`,
                "text": `${ name } : ${ value }`
            });

        
        wrapper.append(label, colorbox, display);
        return wrapper;
    }

    private CreateElement(name: string, opt: {} = {}): any {
        const elem = document.createElement(name);

        for (const attr in opt) {
            if (Object.prototype.hasOwnProperty.call(opt, attr)) {

                switch (attr) {
                    case "class":
                        elem.className = opt[attr];
                        break;
                    case "text":
                        elem.innerText = opt[attr];
                        break;
                    case "html":
                        elem.innerHTML = opt[attr];
                        break;
                
                    default:
                        elem[attr] = opt[attr];
                        break;
                }
            }
        }

        return elem;
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                                PropertyPaneLabel("description", { "text": strings.DescriptionFieldLabel })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
