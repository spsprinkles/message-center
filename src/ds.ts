import { List } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * List Item
 * Add your custom fields here
 */
export interface IListItem extends Types.SP.ListItem {
    Category: string;
    Content: string;
    Message: string;
    MessageId: string;
    Platform: string;
    RoadMapId: string;
    Services: { results: string[]; }
    Severity: string;
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
    // Initializes the application
    static init(): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            return Promise.all([
                // Load the security
                //Security.init(),

                // Load the data
                this.load()
            ]).then(resolve, reject);
        });
    }

    // Loads the list data
    private static _list: List<IListItem>;
    static get List(): List<IListItem> { return this._list; }
    static load(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Initialize the list
            this._list = new List<IListItem>({
                listName: Strings.Lists.Main,
                webUrl: Strings.SourceUrl,
                itemQuery: {
                    GetAllItems: true,
                    OrderBy: ["MessageId"],
                    Top: 5000
                },
                onInitError: reject,
                onInitialized: () => {
                    // Load the filters
                    this.loadFilters();

                    // Resolve the request
                    resolve();
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
            let services = item.Services.results;
            for (let i = 0; i < services.length; i++) {
                filters.services[services[i]] = true;
            }

            // Parse the tags
            let tags = item.Tags.results;
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
            // Refresh the data
            DataSource.List.refresh().then(resolve, reject);
        });
    }
}