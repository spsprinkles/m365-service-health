import { List } from "dattatable";
import { Components, Types, Web } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * List Item
 * Add your custom fields here
 */
export interface IListItem extends Types.SP.ListItem {
    Status: string;
}

/**
 * Data Source
 */
export class DataSource {
    // List
    private static _list: List<IListItem> = null;
    static get List(): List<IListItem> { return this._list; }

    // List Items
    static get ListItems(): IListItem[] { return this.List.Items; }

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

    // Initializes the application
    static init(): PromiseLike<void> {
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
                onInitialized: () => {
                    // Load the status filters
                    this.loadStatusFilters().then(() => {
                        // Resolve the request
                        resolve();
                    }, reject);
                }
            });
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