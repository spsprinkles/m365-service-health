import { List } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import { Security } from "./security";
import Strings from "./strings";

/**
 * List Item
 * Add your custom fields here
 */
export interface IListItem extends Types.SP.ListItem {
    ServiceId: string;
    ServiceIssues: string;
    ServiceStatus: string;
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
                Security.init(),
                // Load the data
                this.load(),
                // Load the Status Filters
                this.loadStatusFilters()
            ]).then(resolve, reject);
        });
    }

    // Loads the list data
    private static _list: List<IListItem> = null;
    static get List(): List<IListItem> { return this._list; }
    static get ListItems(): IListItem[] { return this.List.Items; }
    static load(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Initialize the list
            this._list = new List<IListItem>({
                listName: Strings.Lists.Main,
                itemQuery: {
                    GetAllItems: true,
                    OrderBy: ["Title"],
                    Top: 5000
                },
                onInitError: reject,
                onInitialized: resolve
            });
        });
    }

    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = null;
    static get StatusFilters(): Components.ICheckboxGroupItem[] { return this._statusFilters; }
    static loadStatusFilters(): PromiseLike<Components.ICheckboxGroupItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the status field choices
            let field = this.List.getField("ServiceStatus") as Types.SP.FieldChoice;
            if (field) {
                let items: Components.ICheckboxGroupItem[] = [];

                // Parse the choices
                for (let i = 0; i < field.Choices.results.length; i++) {
                    // Add an item
                    items.push({
                        label: field.Choices.results[i],
                        type: Components.CheckboxGroupTypes.Switch
                    });
                }

                // Set the filters and resolve the promise
                this._statusFilters = items;
                resolve(items);
            } else {
                // Reject the request
                reject("The status field is missing.");
            }
        });
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