import { Dashboard, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
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
            tiles: {
                items: DataSource.ListItems,
                bodyField: "ServiceId",
                filterField: "ServiceStatus",
                subTitleField: "ServiceStatus",
                titleField: "Title",
                onBodyRender: (el, item: IListItem) => {
                    // See if issues exist
                    let issues = item.ServiceIssues ? JSON.parse(item.ServiceIssues) : null;
                    if (issues && issues.length > 0) {
                        // Render a button to show the issues
                        Components.Tooltip({
                            el,
                            content: "Click to view the issues with this service.",
                            btnProps: {
                                text: "Issues",
                                type: Components.ButtonTypes.OutlineWarning,
                                onClick: () => {
                                    // Set the modal header
                                    Modal.setHeader(item.Title + " Issues");

                                    // Parse the items
                                    let items: Components.IListGroupItem[] = [];
                                    for(let i=0; i<issues.length; i++) {
                                        let issue = issues[i];

                                        // Add the issue
                                        items.push({
                                            content: `
                                                <b>Title:</b> ${issue.title}<br/>
                                                <b>Feature:</b> ${issue.feature}<br/>
                                                <b>Description:</b> ${issue.impactDescription}
                                            `
                                        });
                                    }

                                    // Set the body
                                    Components.ListGroup({
                                        el: Modal.BodyElement,
                                        items
                                    });

                                    // Show the modal
                                    Modal.show();
                                }
                            }
                        });
                    }
                }
            },
            footer: {
                itemsEnd: [
                    {
                        text: "v" + Strings.Version
                    }
                ]
            }
        });
    }
}