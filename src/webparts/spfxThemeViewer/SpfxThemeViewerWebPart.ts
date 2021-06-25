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
            msg = this.CreateElement("div", { "text": "From: window.__themeState__.theme" });

        for (const name in window.__themeState__.theme) {
            if (Object.prototype.hasOwnProperty.call(window.__themeState__.theme, name)) {
                const color = window.__themeState__.theme[name];
                root.append(this.CreateThemeBox(name, color));
            }
        }

        this.domElement.append(msg, root);
    }

    private CreateThemeBox(name: string, color: string): HTMLDivElement {
        const wrapper = this.CreateElement("div", { "class": styles.boxWrapper }),
            label = this.CreateElement("div", { "text": name, "title": color }),
            colorbox = this.CreateElement("div", {
                "style": `background-color: ${ color }`,
                "title": `Click to display value in the console.\n${ name }: ${ color }`,
                "class": styles.box,
                "onclick": evt => {
                    console.info(name, color);
                }
            });

        
        wrapper.append(label, colorbox);
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
