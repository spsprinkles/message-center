import { List } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import { convertCAML } from "./common";
import { Security } from "./security";
import Strings from "./strings";

/**
 * List Item
 * Add your custom fields here
 */
export interface IListItem extends Types.SP.ListItem {
    Category: string;
    Content: string;
    IsApproved: boolean;
    IsMajorChange: boolean;
    Message: string;
    MessageId: string;
    Notes: string;
    Platform: string;
    PublishedDate: string;
    RoadMapId: string;
    Services: { results: string[]; }
    Severity: string;
    Status: string;
    Summary: string;
    Tags: { results: string[]; }
}

/**
 * Message
 */
export interface IMessage {
    startDateTime: string;
    endDateTime: string;
    lastModifiedDateTime: string;
    title: string;
    id: string;
    category: string;
    severity: string;
    tags: string[];
    isMajorChange: boolean;
    actionRequiredByDateTime: string;
    services: string[];
    hasAttachments: boolean;
    viewPoint: string;
    details: {
        name: string;
        value: string;
    }[];
    body: {
        contentType: string;
        content: string;
    }
}

/**
 * Data Source
 */
export class DataSource {
    // Gets the item id from the query string
    static getItemIdFromQS() {
        // Get the id from the querystring
        let qs = document.location.search.split('?');
        qs = qs.length > 1 ? qs[1].split('&') : [];
        for (let i = 0; i < qs.length; i++) {
            let qsItem = qs[i].split('=');
            let key = qsItem[0];
            let value = qsItem[1];

            // See if this is the "id" key
            if (key == "item-id") {
                // Return the item
                return parseInt(value);
            }
        }
    }

    // Initializes the application
    static init(): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            return Promise.all([
                // Load the security
                Security.init(),

                // Load the data
                this.load()
            ]).then(resolve, reject);
        });
    }

    // Loads the list data
    private static _items: { [key: string]: IListItem[] } = null;
    static get ApprovedItems(): IListItem[] { return this._items.Approved; }
    static get NeedsApprovalItems(): IListItem[] { return this._items.NeedsApproval; }
    static get NotPublishedItems(): IListItem[] { return this._items.NotPublished; }
    private static _list: List<IListItem>;
    static get List(): List<IListItem> { return this._list; }
    static load(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Clear the items
            this._items = {
                Approved: [],
                NeedsApproval: [],
                NotPublished: []
            }

            // Initialize the list
            this._list = new List<IListItem>({
                listName: Strings.Lists.Main,
                webUrl: Strings.SourceUrl,
                itemQuery: {
                    GetAllItems: true,
                    OrderBy: ["MessageId desc"],
                    Top: 5000
                },
                onInitError: reject,
                onInitialized: () => {
                    // Load the filters
                    this.loadFilters();

                    // Load the status items
                    this.loadStatus();

                    // Resolve the request
                    resolve();
                },
                onItemLoading: item => {
                    // Convert the category & severity
                    item.Category = convertCAML(item.Category);
                    item.Severity = convertCAML(item.Severity);

                    // See if this item is approved
                    if (item.IsApproved) {
                        this._items.Approved.push(item);
                    }
                    // Else, see if there is a status set
                    else if (item.Status) {
                        this._items.NotPublished.push(item);
                    }
                    // Else, this item hasn't been reviewed
                    else {
                        this._items.NeedsApproval.push(item);
                    }
                }
            });
        });
    }

    // Filters
    private static _filters: { [key: string]: Components.ICheckboxGroupItem[] } = null;
    static get CategoryFilters(): Components.ICheckboxGroupItem[] { return this._filters.category; }
    static get SeverityFilters(): Components.ICheckboxGroupItem[] { return this._filters.severity; }
    static get ServicesFilters(): Components.ICheckboxGroupItem[] { return this._filters.services; }
    static get TagsFilters(): Components.ICheckboxGroupItem[] { return this._filters.tags; }
    static loadFilters() {
        let filters = {
            category: {},
            services: {},
            severity: {},
            tags: {}
        }

        // Clear the filters
        this._filters = {
            category: [],
            services: [],
            severity: [],
            tags: []
        }

        // Parse the items
        for (let i = 0; i < this.List.Items.length; i++) {
            let item = this.List.Items[i];

            // Set the filter values
            item.Category ? filters.category[item.Category] = true : null;
            item.Severity ? filters.severity[item.Severity] = true : null;

            // Parse the services
            let services = item.Services?.results || [];
            for (let i = 0; i < services.length; i++) {
                filters.services[services[i]] = true;
            }

            // Parse the tags
            let tags = item.Tags?.results || [];
            for (let i = 0; i < tags.length; i++) {
                filters.tags[tags[i]] = true;
            }
        }

        // Set the filters
        let setFilters = (key: string) => {
            // Parse the filter values
            for (let filterValue in filters[key]) {
                // Add the filter value
                this._filters[key].push({
                    data: filterValue,
                    label: filterValue,
                    type: Components.CheckboxGroupTypes.Switch
                });
            }

            // Sorty the filters
            this._filters[key] = this._filters[key].sort((a, b) => {
                if (a.label < b.label) { return -1; }
                if (a.label > b.label) { return 1; }
                return 0;
            });
        }

        // Set the filters
        setFilters("category");
        setFilters("services");
        setFilters("severity");
        setFilters("tags");
    }

    // Refreshes the list data
    static refresh(): PromiseLike<IListItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Clear the items
            this._items = {
                Approved: [],
                NeedsApproval: [],
                NotPublished: []
            }

            // Refresh the data
            DataSource.List.refresh().then(resolve, reject);
        });
    }

    // Status choice options
    private static _statusItems: Components.IDropdownItem[] = null;
    static get StatusItems(): Components.IDropdownItem[] { return this._statusItems; }
    private static loadStatus() {
        // Get the status field
        let fldStatus = this.List.getField("Status") as Types.SP.FieldChoice;

        // Clear the items
        this._statusItems = [];

        // Parse the choices
        for (let i = 0; i < fldStatus.Choices.results.length; i++) {
            // Append the choice
            this._statusItems.push({
                text: fldStatus.Choices.results[i],
                value: fldStatus.Choices.results[i]
            });
        }
    }
}