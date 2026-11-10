import { List } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import * as moment from 'moment-timezone';
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
    PublishedDateAsNumber: number;
    RoadMapId: string;
    Services: { results: string[]; }
    ServicesAsString: string;
    Severity: string;
    State: "Approved" | "Needs Approval" | "Not Published" | "";
    Status: string;
    Summary: string;
    Tags: { results: string[]; }
    TagsAsString: string;
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
            Promise.all([
                // Load the data
                DataSource.load(),

                // Load the security
                Security.init()
            ]).then(resolve, reject);
        });
    }

    // Filters
    private static _services: string[] = null;
    private static _tags: string[] = null;
    static get Services(): string[] { return this._services; }
    static get Tags(): string[] { return this._tags; }

    // Loads the list data
    private static _items: IListItem[] = null;
    static get Items(): IListItem[] { return this._items; }
    private static _list: List<IListItem>;
    static get List(): List<IListItem> { return this._list; }
    static load(onItemsLoaded?: (items: IListItem[]) => void): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Clear the items
            this._items = [];

            // The items to send to the event
            let items: IListItem[] = [];

            // Set the filters
            let services = {};
            let tags = {};

            // Adds the item to the array
            let loopId = null;
            let addItem = (item: IListItem) => {
                // Append the item
                items.push(item);
                this._items.push(item);

                // See if there is a loop
                if (loopId) { return; }

                // Set the loop id and wait 100 ms
                loopId = setTimeout(() => {
                    // Call the event
                    onItemsLoaded ? onItemsLoaded(items) : null;

                    // Clear the items and loop id
                    items = [];
                    loopId = null;
                }, 100);
            }

            // Initialize the list
            this._list = new List<IListItem>({
                listName: Strings.Lists.Main,
                webUrl: Strings.SourceUrl,
                itemQuery: {
                    GetAllItems: true,
                    OrderBy: [Strings.SortField + " desc"],
                    Top: 5000
                },
                onInitError: reject,
                onInitialized: () => {
                    // Set the services and tags
                    this._services = Object.keys(services).sort();
                    this._tags = Object.keys(tags).sort();

                    // Load the status items
                    this.loadStatus();

                    // Resolve the request
                    resolve();
                },
                onItemLoading: item => {
                    // Convert the category & severity
                    item.Category = convertCAML(item.Category);
                    item.Severity = convertCAML(item.Severity);

                    // Set the string values
                    item.PublishedDateAsNumber = item.PublishedDate ? (new Date(item.PublishedDate)).getTime() : 0;
                    item.ServicesAsString = item.Services?.results?.join("|") || "";
                    item.TagsAsString = item.Tags?.results?.join("|") || "";

                    // Parse the services and tags
                    (item.Services?.results || []).forEach(service => { services[service] = true; });
                    (item.Tags?.results || []).forEach(tag => { tags[tag] = true; });

                    // See if this item is approved
                    if (item.IsApproved) {
                        // Set the state
                        item.State = "Approved";
                    }
                    // Else, see if there is a status set
                    else if (item.Status) {
                        item.State = "Not Published";
                    }
                    // Else, this item hasn't been reviewed
                    else {
                        item.State = "Needs Approval";
                    }

                    // Add the item if the user is an admin
                    if (Security.IsAdmin) {
                        // Append the item
                        addItem(item);
                    }
                    // Else, only show approved items
                    else if (item.State != "Approved") {
                        // Append the item
                        addItem(item);
                    }
                }
            });
        });
    }

    // Refreshes the list data
    static refresh(): PromiseLike<IListItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Clear the items
            this._items = [];

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