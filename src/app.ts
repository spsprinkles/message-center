import { Dashboard } from "dattatable";
import { Components, CustomIcons, CustomIconTypes } from "gd-sprest-bs";
import { DataSource, IListItem } from "./ds";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    // Constructor
    constructor(el: HTMLElement) {
        // Render the dashboard
        this.render(el);
    }

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Create the dashboard
        let dashboard = new Dashboard({
            el,
            hideHeader: true,
            useModal: true,
            filters: {
                items: [
                    {
                        header: "By Category",
                        items: DataSource.CategoryFilters,
                        multi: true,
                        onFilter: (values: string[]) => {
                            // Filter the tiles
                            dashboard.filterTiles(values);
                        }
                    },
                    {
                        header: "By Services",
                        items: DataSource.ServicesFilters,
                        multi: true,
                        onFilter: (values: string[]) => {
                            // Filter the tiles
                            dashboard.filterTiles(values);
                        }
                    },
                    {
                        header: "By Severity",
                        items: DataSource.SeverityFilters,
                        multi: true,
                        onFilter: (values: string[]) => {
                            // Filter the tiles
                            dashboard.filterTiles(values);
                        }
                    },
                    {
                        header: "By Tags",
                        items: DataSource.TagsFilters,
                        multi: true,
                        onFilter: (values: string[]) => {
                            // Filter the tiles
                            dashboard.filterTiles(values);
                        }
                    }
                ]
            },
            navigation: {
                title: Strings.ProjectName
            },
            footer: {
                itemsEnd: [
                    {
                        text: "v" + Strings.Version
                    }
                ]
            },
            tiles: {
                items: DataSource.List.Items,
                colSize: 3,
                paginationLimit: 9,
                showFooter: false,
                filterFields: ["Category", "Services", "Severity", "Tags"],
                titleFields: ["MessageId"],
                subTitleFields: ["Title"],
                bodyFields: ["Summary"],
                onHeaderRendered: (el, item: IListItem) => {
                    el.classList.add("text-center");

                    // Render the icons
                    this.renderIcons(el, item);
                }
            }
        });
    }

    // Renders the icons for the item
    private renderIcons(el: HTMLElement, item: IListItem) {
        // Define the height/width of the icons
        let height = 32;
        let width = 32;

        // Parse the services
        let elIconTypes = [];
        for (let i = 0; i < item.Services.results.length; i++) {
            let elIconType = null;
            let service = item.Services.results[i];

            switch (service) {
                case "Azure Information Protection":
                    elIconType = CustomIconTypes.aIP;
                    break;

                case "Exchange Online":
                    elIconType = CustomIconTypes.exchange;
                    break;

                case "Microsoft 365 apps":
                case "Microsoft 365 for the web":
                case "Microsoft 365 suite":
                    elIconType = CustomIconTypes.m365;
                    break;

                case "Microsoft Entra":
                    elIconType = CustomIconTypes.entra;
                    break;

                case "Microsoft Forms":
                    elIconType = CustomIconTypes.forms;
                    break;

                case "Microsoft Stream":
                    elIconType = CustomIconTypes.stream;
                    break;

                case "Microsoft Teams":
                    elIconType = CustomIconTypes.teams;
                    break;

                case "Microsoft Viva":
                    elIconType = CustomIconTypes.viva;
                    break;

                case "SharePoint Online":
                    elIconType = CustomIconTypes.sharePoint;
                    break;

                // Default Icon
                default:
                    break;
            }

            // See if this icon hasn't been rendered
            if (elIconType && elIconTypes.indexOf(elIconType) < 0) {
                // Create the icon
                let elIcon = CustomIcons(elIconType, height, width);
                elIcon.style.height = height.toString();
                elIcon.style.width = width.toString();
                elIcon.setAttribute("aria-hidden", "true");
                elIcon.style.pointerEvents = "none";
                elIcon.setAttribute("focusable", "false");

                // Append the type
                elIconTypes.push(elIconType);

                // Add a tooltip
                Components.Tooltip({
                    content: service,
                    type: Components.TooltipTypes.LightBorder,
                    target: elIcon as any
                });

                // Append the icon
                el.appendChild(elIcon);
            }
        }
    }
}