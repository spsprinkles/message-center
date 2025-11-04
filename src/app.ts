import { CanvasForm, Dashboard, LoadingDialog, Modal } from "dattatable";
import { Components, CustomIcons, CustomIconTypes } from "gd-sprest-bs";
import { clipboard } from "gd-sprest-bs/build/icons/svgs/clipboard";
import { funnel } from "gd-sprest-bs/build/icons/svgs/funnel";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import * as moment from 'moment-timezone';
import { DataSource, IListItem } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
import Strings from "./strings";

// Items to view
enum ItemsToShow {
    ShowAll = 0,
    Approved = 1,
    NeedsApproval = 2,
    NotPublished = 3
}


/**
 * Main Application
 */
export class App {
    private _dashboard: Dashboard = null;
    private _currentView: number = ItemsToShow.ShowAll;
    private _elNavigation: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement, itemId?: number) {
        // Render the dashboard
        this.render(el);

        // See if we are displaying an item
        let item = itemId > 0 ? DataSource.List.getItem(itemId) : null;
        if (item) {
            // Show the item
            this.showMoreInfo(item);
        }
    }

    // Refreshes the dashboard
    private refresh(view: ItemsToShow) {
        // Set the current view
        this._currentView = view;

        // See what items to show
        switch (view) {
            case ItemsToShow.Approved:
                // Show the approved items
                this._dashboard.refresh(DataSource.ApprovedItems);
                break;
            case ItemsToShow.NotPublished:
                // Show the items that have been reviewed
                this._dashboard.refresh(DataSource.NotPublishedItems);
                break;
            case ItemsToShow.NeedsApproval:
                // Show the items that need to be reviewed
                this._dashboard.refresh(DataSource.NeedsApprovalItems);
                break;
            default:
                // See if this is an admin
                if (Security.IsAdmin) {
                    // Show all the items
                    this._dashboard.refresh(DataSource.List.Items);
                } else {
                    // Show only the approved items
                    this._dashboard.refresh(DataSource.ApprovedItems);
                }
                break;
        }
    }

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Create the dashboard
        this._dashboard = new Dashboard({
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
                            this._dashboard.filterTiles(values);
                        }
                    },
                    {
                        header: "By Services",
                        items: DataSource.ServicesFilters,
                        multi: true,
                        onFilter: (values: string[]) => {
                            // Filter the tiles
                            this._dashboard.filterTiles(values);
                        }
                    },
                    {
                        header: "By Severity",
                        items: DataSource.SeverityFilters,
                        multi: true,
                        onFilter: (values: string[]) => {
                            // Filter the tiles
                            this._dashboard.filterTiles(values);
                        }
                    },
                    {
                        header: "By Tags",
                        items: DataSource.TagsFilters,
                        multi: true,
                        onFilter: (values: string[]) => {
                            // Filter the tiles
                            this._dashboard.filterTiles(values);
                        }
                    }
                ]
            },
            navigation: {
                title: Strings.ProjectName,
                itemsEnd: Security.IsAdmin ? [
                    {
                        isButton: true,
                        className: "btn-icon btn-outline-light me-2 p-2 py-1",
                        iconSize: 22,
                        iconType: gearWideConnected,
                        text: "Settings",
                        items: [
                            {
                                text: "App Settings",
                                onClick: () => {
                                    // Show the install modal
                                    InstallationModal.show(true);
                                }
                            },
                            {
                                text: Strings.Lists.Main + " List",
                                onClick: () => {
                                    // Show the FAQ list in a new tab
                                    window.open(Strings.SourceUrl + "/Lists/" + Strings.Lists.Main.split(" ").join(""), "_blank");
                                }
                            },
                            {
                                text: Security.AdminGroup.Title + " Group",
                                onClick: () => {
                                    // Show the settings in a new tab
                                    window.open(Strings.SourceUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.AdminGroup.Id);
                                }
                            },
                            {
                                text: Security.MemberGroup.Title + " Group",
                                onClick: () => {
                                    // Show the settings in a new tab
                                    window.open(Strings.SourceUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.MemberGroup.Id);
                                }
                            },
                            {
                                text: Security.VisitorGroup.Title + " Group",
                                onClick: () => {
                                    // Show the settings in a new tab
                                    window.open(Strings.SourceUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.VisitorGroup.Id);
                                }
                            }
                        ]
                    },
                    {
                        isButton: true,
                        className: "btn-icon btn-outline-light me-2 p-2 py-1",
                        iconSize: 22,
                        iconType: funnel,
                        text: "View Items",
                        items: [
                            {
                                text: "Show All",
                                onClick: (item, ev) => {
                                    // Refresh the items
                                    this.refresh(ItemsToShow.ShowAll);
                                }
                            },
                            {
                                text: "Approved",
                                onClick: (item, ev) => {
                                    // Refresh the items
                                    this.refresh(ItemsToShow.Approved);
                                }
                            },
                            {
                                text: "Needs Approval",
                                onClick: (item, ev) => {
                                    // Refresh the items
                                    this.refresh(ItemsToShow.NeedsApproval);
                                }
                            },
                            {
                                text: "Not Published",
                                onClick: (item, ev) => {
                                    // Refresh the items
                                    this.refresh(ItemsToShow.NotPublished);
                                }
                            }
                        ]
                    }
                ] : null,
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
                },
                onRendered: el => {
                    // Save a reference to the navigation
                    this._elNavigation = el;
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
                items: Security.IsAdmin ? DataSource.List.Items : DataSource.ApprovedItems,
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

                    // Render the review button
                    this.renderReviewButton(el, item);

                    // Render the link button
                    this.renderLinkButton(el, item);
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
                    type: Components.ButtonTypes.OutlinePrimary,
                    isSmall: true,
                    onClick: () => {
                        // Show the information for this item
                        this.showMoreInfo(item);
                    }
                }
            });
        }
    }

    // Renders the link button
    private renderLinkButton(el: HTMLElement, item: IListItem) {
        // Add a button for additional details
        Components.Tooltip({
            el,
            content: "Generates a link to this item.",
            btnProps: {
                className: "ms-2",
                text: "Link",
                type: Components.ButtonTypes.OutlinePrimary,
                isSmall: true,
                onClick: () => {
                    // Clear the modal
                    Modal.clear();

                    // Set the header
                    Modal.setHeader("Generate Link");

                    // Generate the url for this item
                    let linkUrl = `${document.location.origin}${document.location.pathname}?item-id=${item.Id}`;

                    // Copy it to the clipboard
                    navigator.clipboard.writeText(linkUrl);

                    // Set the body
                    Components.Form({
                        el: Modal.BodyElement,
                        controls: [
                            {
                                label: "Item Url:",
                                type: Components.FormControlTypes.Readonly,
                                description: "The url has been automatically copied to your clipboard.",
                                value: linkUrl
                            }
                        ]
                    });

                    // Render the close button
                    Components.Button({
                        el: Modal.FooterElement,
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Close the modal
                            Modal.hide();
                        }
                    });

                    // Show the form
                    Modal.show();
                }
            }
        });
    }

    // Renders the review button
    private renderReviewButton(el: HTMLElement, item: IListItem) {
        // Ensure this is an admin
        if (!Security.IsAdmin) { return; }

        // See if this item is approved
        if (item.IsApproved) { return; }

        // Add a button for additional details
        Components.Tooltip({
            el,
            content: "Click to review the item for visibility.",
            btnProps: {
                className: "ms-2",
                text: "Review",
                type: Components.ButtonTypes.OutlinePrimary,
                isSmall: true,
                onClick: () => {
                    // Display an approval modal
                    Modal.clear();
                    Modal.setHeader("Review Item");

                    // Render the form
                    let form = Components.Form({
                        el: Modal.BodyElement,
                        value: item,
                        controls: [
                            {
                                name: "Approve",
                                label: "Approve",
                                type: Components.FormControlTypes.Switch,
                                description: "Approves the item for all users to see.",
                                value: false,
                                onChange: (item) => {
                                    if (item) {
                                        form.getControl("Status").hide();
                                        form.getControl("Notes").hide();
                                        btnSubmit.setText("Approve");
                                    } else {
                                        form.getControl("Status").show();
                                        form.getControl("Notes").show();
                                        btnSubmit.setText("Update");
                                    }
                                }
                            } as Components.IFormControlPropsSwitch,
                            {
                                name: "Status",
                                label: "Status",
                                type: Components.FormControlTypes.Dropdown,
                                description: "The reason why the item is not being published.",
                                items: DataSource.StatusItems
                            } as Components.IFormControlPropsDropdown,
                            {
                                name: "Notes",
                                label: "Notes",
                                description: "Details for why an item should not be published.",
                                type: Components.FormControlTypes.TextArea
                            }
                        ]
                    });

                    // Render the footer
                    let btnSubmit: Components.IButton = null;
                    Components.TooltipGroup({
                        el: Modal.FooterElement,
                        tooltips: [
                            {
                                content: "Click to approve the item and make it available to all users.",
                                btnProps: {
                                    assignTo: btn => { btnSubmit = btn; },
                                    text: "Update",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Hide the modal
                                        Modal.hide();

                                        // Show a loading dialog
                                        LoadingDialog.setHeader("Saving Item");
                                        LoadingDialog.setBody("This will close after the item has been updated...");
                                        LoadingDialog.show();

                                        // Get the form values and see if we are approving the item
                                        let values = form.getValues();
                                        if (values["Approve"]) {
                                            // Update the item
                                            item.update({
                                                IsApproved: true,
                                                Status: "",
                                                Notes: ""
                                            }).execute(() => {
                                                // Refresh the data source
                                                DataSource.refresh().then(() => {
                                                    // Refresh the app
                                                    this.refresh(this._currentView);

                                                    // Close the loading dialog
                                                    LoadingDialog.hide();
                                                });
                                            });
                                        } else {
                                            // Update the item
                                            item.update({
                                                Status: values["Status"].text,
                                                Notes: values["Notes"]
                                            }).execute(() => {
                                                // Refresh the data source
                                                DataSource.refresh().then(() => {
                                                    // Refresh the app
                                                    this.refresh(this._currentView);

                                                    // Close the loading dialog
                                                    LoadingDialog.hide();
                                                });
                                            });
                                        }
                                    }
                                }
                            }
                        ]
                    });

                    // Show the modal
                    Modal.show();
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

    // Shows more information on the item
    private showMoreInfo(item: IListItem) {
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

        // Set the roadmap links
        let roadmapLinks = "";
        let roadmapIds = (item.RoadMapId || "").split(',');
        for (let i = 0; i < roadmapIds.length; i++) {
            // Ensure the id exists
            let roadmapId = roadmapIds[i].trim();
            if (roadmapId.length > 0) {
                // Add a space for multiple links
                i > 0 ? roadmapLinks += " " : null;

                // Add the link
                roadmapLinks += `<a target="_blank" href="https://www.microsoft.com/en-US/microsoft-365/roadmap?filters=&searchterms=${roadmapId}">${roadmapId}</a>`;
            }
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
                    <div class="mb-3">${roadmapLinks}</div>
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