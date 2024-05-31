import { CanvasForm, Dashboard } from "dattatable";
import { Components, CustomIcons, CustomIconTypes } from "gd-sprest-bs";
import { clipboard } from "gd-sprest-bs/build/icons/svgs/clipboard";
import * as moment from 'moment-timezone';
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
                title: Strings.ProjectName,
                onRendering: props => {
                    // Update the navigation properties
                    props.className = "navbar-expand rounded-top";
                    props.type = Components.NavbarTypes.Primary;

                    // Add a logo to the navbar brand
                    let navText = document.createElement("div");
                    navText.classList.add("ms-2");
                    navText.append(Strings.ProjectName);
                    let navIcon = document.createElement("div");
                    navIcon.classList.add("d-flex");
                    navIcon.appendChild(clipboard(32, 32));
                    navIcon.appendChild(navText);
                    props.brand = navIcon;
                }
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
                colSize: Strings.TileColumnSize,
                paginationLimit: Strings.TilePageSize,
                filterFields: ["Category", "Services", "Severity", "Tags"],
                titleFields: ["MessageId", "RoadMapId"],
                titleTemplate: `
                    <div class="d-flex justify-content-between align-items-center">
                        <span>{MessageId}</span>
                        <span>{RoadMapId}</span>
                    </div>
                `,
                onTitleRendered: (el) => {
                    // Update the class names
                    el.classList.remove("h5");
                    el.classList.add("fs-6");
                },
                subTitleFields: ["Title"],
                onSubTitleRendered: (el) => {
                    // Add spacing
                    el.classList.add("mt-1");
                    el.classList.add("mb-1");
                },
                bodyFields: ["Summary"],
                onBodyRendered: (el) => {
                    // Update the padding for the top
                    el.classList.add("pt-1");
                },
                onHeaderRendered: (el, item: IListItem) => {
                    // Add the classes
                    el.classList.add("d-flex");
                    el.classList.add("justify-content-between");
                    el.classList.add("align-items-center");

                    // Render the header template
                    el.innerHTML = `
                        <span class="h6 m-0">Message ID</span>
                        <span class="card-icons"></span>
                        <span class="h6 m-0">Road Map ID</span>
                    `;

                    // Render the icons
                    this.renderIcons(el.querySelector(".card-icons"), item);
                },
                onFooterRendered: (el, item) => {
                    // Center the button
                    el.classList.add("text-center");

                    // Render the details button
                    this.renderDetailsButton(el, item);
                },
                onColumnRendered: (el) => {
                    // Add spacing for the top of the card
                    el.classList.add("mt-2");
                },
                onCardRendered: (el) => {
                    // Add the class names for making the heights match
                    el.classList.add("h-100");
                }
            }
        });
    }

    // Renders the details button
    private renderDetailsButton(el: HTMLElement, item: IListItem) {
        // Ensure content exists
        if (item.Content) {
            // Add a button for additional details
            Components.Tooltip({
                el,
                content: Strings.MoreInfoTooltip,
                btnProps: {
                    text: Strings.MoreInfo,
                    type: Components.ButtonTypes.OutlineSuccess,
                    isSmall: true,
                    onClick: () => {
                        // Clear the canvas
                        CanvasForm.clear();

                        // Set the properties
                        CanvasForm.setSize(Components.OffcanvasSize.Medium3);
                        CanvasForm.setType(Components.OffcanvasTypes.End);

                        // Set the header
                        CanvasForm.setHeader(item.Title);

                        // Make the headers bold
                        let content = item.Content
                            .replace(/\[/g, "<b>").replace(/\]/g, "</b>");

                        // Parse the tags
                        let badges = [];
                        let tags = item.Tags?.results || [];
                        for (let i = 0; i < tags.length; i++) {
                            badges.push(Components.Badge({
                                className: "text-bg-secondary",
                                content: tags[i]
                            }).el.outerHTML);
                        }

                        // Get the icons
                        let el = document.createElement("div");
                        this.renderIcons(el, item);

                        // Create the summary
                        let summary = Components.Alert({
                            header: "Summary",
                            content: item.Summary,
                            type: Components.AlertTypes.Secondary
                        });

                        // Set the published date
                        let publishedDate = "";
                        if (item.PublishedDate) {
                            // Format the date/time
                            publishedDate = moment.utc(item.PublishedDate).tz(Strings.TimeZone).format(Strings.TimeFormat);
                        }

                        // Render the body
                        CanvasForm.BodyElement.innerHTML = `
                            <div class="row">
                                <div class="col-9">
                                    ${summary.el.outerHTML}
                                    ${content}
                                </div>
                                <div class="col-3">
                                    <div class="fs-6">Service:</div>
                                    <div class="mb-3">${el.innerHTML}</div>
                                    <div class="fs-6">Platform:</div>
                                    <div class="mb-3">${item.Platform || ""}</div>
                                    <div class="fs-6">Message ID:</div>
                                    <div class="mb-3 fw-semibold">${item.MessageId || ""}</div>
                                    <div class="fs-6">Roadmap ID:</div>
                                    <div class="mb-3"><a target="_blank" href="https://www.microsoft.com/en-US/microsoft-365/roadmap?filters=&searchterms=${item.RoadMapId}">${item.RoadMapId || ""}</a></div>
                                    <div class="fs-6">Published:</div>
                                    <div class="mb-3">${publishedDate}</div>
                                    <div class="fs-6">Tag(s):</div>
                                    <div class="mb-3">${badges.join('\n')}</div>
                                </div>
                            </div>
                            <div class="footer d-flex justify-content-end">
                            </div>
                        `;

                        // Render the close button
                        Components.Button({
                            el: CanvasForm.BodyElement.querySelector(".footer"),
                            text: "Close",
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                // Close the canvas
                                CanvasForm.hide();
                            }
                        });

                        // Show the form
                        CanvasForm.show();
                    }
                }
            });
        }
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

                case "Dynamics 365 Apps":
                    elIconType = CustomIconTypes.dynamics;
                    break;

                case "Exchange Online":
                    elIconType = CustomIconTypes.exchange;
                    break;

                case "Microsoft 365 apps":
                case "Microsoft 365 for the web":
                case "Microsoft 365 suite":
                    elIconType = CustomIconTypes.m365;
                    break;

                case "Microsoft Dataverse":
                    elIconType = CustomIconTypes.dataverse;
                    break;

                case "Microsoft Defender":
                case "Microsoft Defender XDR":
                case "Microsoft Defender for Cloud Apps":
                    elIconType = CustomIconTypes.defender;
                    break;

                case "Microsoft Entra":
                    elIconType = CustomIconTypes.entra;
                    break;

                case "Microsoft Forms":
                    elIconType = CustomIconTypes.forms;
                    break;

                case "Microsoft Intune":
                    elIconType = CustomIconTypes.intune;
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

                case "OneDrive for Business":
                    elIconType = CustomIconTypes.oneDrive;
                    break;

                case "Power Apps":
                    elIconType = CustomIconTypes.powerApps;
                    break;

                case "Power Platform":
                    elIconType = CustomIconTypes.powerPlatform;
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