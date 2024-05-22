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
                    //this.loadFilters();

                    // Resolve the request
                    resolve();
                }
            });
        });
    }

    /*
    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = null;
    static get StatusFilters(): Components.ICheckboxGroupItem[] { return this._statusFilters; }
    static loadStatusFilters() {
        // Get the status field choices
        let field = this.List.getField("ServiceStatus") as Types.SP.FieldChoice;
        if (field) {
            let items: Components.ICheckboxGroupItem[] = [];

            // Parse the choices
            for (let i = 0; i < field.Choices.results.length; i++) {
                // Add an item
                items.push({
                    data: field.Choices.results[i],
                    label: getStatusTitle(field.Choices.results[i]),
                    type: Components.CheckboxGroupTypes.Switch
                });
            }

            // Set the filters and resolve the promise
            this._statusFilters = items;
        }
    }
    */

    // Refreshes the list data
    static refresh(): PromiseLike<IListItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Refresh the data
            DataSource.List.refresh().then(resolve, reject);
        });
    }
}