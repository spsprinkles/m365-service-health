import { Dashboard } from "dattatable";
import { Components } from "gd-sprest-bs";
import * as jQuery from "jquery";
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
                items: [{
                    header: "By Status",
                    items: DataSource.StatusFilters,
                    onFilter: (value: string) => {
                        // Filter the table
                        dashboard.filter(1, value);
                    }
                }]
            },
            navigation: {
                title: Strings.ProjectName,
                items: [
                    {
                        className: "btn-outline-light",
                        text: "Create Item",
                        isButton: true,
                        onClick: () => {
                            // Show the new form
                            DataSource.List.newForm({
                                onUpdate: () => {
                                    // Refresh the data
                                    DataSource.refresh().then(() => {
                                        // Refresh the table
                                        dashboard.refresh(DataSource.ListItems);
                                    });
                                }
                            });
                        }
                    }
                ]
            },
            footer: {
                itemsEnd: [
                    {
                        text: "v" + Strings.Version
                    }
                ]
            },
            table: {
                rows: DataSource.ListItems,
                dtProps: {
                    dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                    columnDefs: [
                        {
                            "targets": 0,
                            "orderable": false,
                            "searchable": false
                        }
                    ],
                    createdRow: function (row, data, index) {
                        jQuery('td', row).addClass('align-middle');
                    },
                    drawCallback: function (settings) {
                        let api = new jQuery.fn.dataTable.Api(settings) as any;
                        let div = api.table().container() as HTMLDivElement;
                        let table = api.table().node() as HTMLTableElement;
                        div.querySelector(".dataTables_info").classList.add("text-center");
                        div.querySelector(".dataTables_length").classList.add("pt-2");
                        div.querySelector(".dataTables_paginate").classList.add("pt-03");
                        table.classList.remove("no-footer");
                        table.classList.add("tbl-footer");
                        table.classList.add("table-striped");
                    },
                    headerCallback: function (thead, data, start, end, display) {
                        jQuery('th', thead).addClass('align-middle');
                    },
                    // Order by the 1st column by default; ascending
                    order: [[1, "asc"]]
                },
                columns: [
                    {
                        name: "",
                        title: "Title",
                        onRenderCell: (el, column, item: IListItem) => {
                            // Render a buttons
                            Components.ButtonGroup({
                                el,
                                buttons: [
                                    {
                                        text: item.Title,
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the display form
                                            DataSource.List.viewForm({
                                                itemId: item.Id
                                            });
                                        }
                                    },
                                    {
                                        text: "Edit",
                                        type: Components.ButtonTypes.OutlineSuccess,
                                        onClick: () => {
                                            // Show the edit form
                                            DataSource.List.editForm({
                                                itemId: item.Id,
                                                onUpdate: () => {
                                                    // Refresh the data
                                                    DataSource.refresh().then(() => {
                                                        // Refresh the table
                                                        dashboard.refresh(DataSource.ListItems);
                                                    });
                                                }
                                            });
                                        }
                                    }
                                ]
                            });
                        }
                    },
                    {
                        name: "Status",
                        title: "Status"
                    }
                ]
            }
        });
    }
}