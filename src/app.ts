import { Dashboard, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { wrenchAdjustableCircle } from "gd-sprest-bs/build/icons/svgs/wrenchAdjustableCircle";
import { DataSource, IListItem } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
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

    // Renders the navigation items
    private generateNavItems() {
        // Create the settings menu items
        let itemsEnd: Components.INavbarItem[] = [];
        if (Security.IsAdmin) {
            itemsEnd.push(
                {
                    className: "btn-icon btn-outline-light me-2 p-2 py-1",
                    text: "Settings",
                    iconSize: 22,
                    iconType: gearWideConnected,
                    isButton: true,
                    items: [
                        {
                            text: "App Settings",
                            onClick: () => {
                                // Show the install modal
                                InstallationModal.show(true);
                            }
                        },
                        {
                            text: Strings.Lists.Main + " list",
                            onClick: () => {
                                // Show the FAQ list in a new tab
                                window.open(Strings.SourceUrl + "/Lists/" + Strings.Lists.Main, "_blank");
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
                }
            );
        }

        // Return the nav items
        return itemsEnd;
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
            footer: {
                onRendering: props => {
                    // Update the properties
                    props.className = "footer p-0";
                },
                onRendered: (el) => {
                    el.querySelector("nav.footer").classList.remove("bg-light");
                    el.querySelector("nav.footer .container-fluid").classList.add("p-0");
                },
                itemsEnd: [{
                    className: "pe-none text-body",
                    text: "v" + Strings.Version,
                    onRender: (el) => {
                        // Hide version footer in a modern page
                        Strings.IsClassic ? null : el.classList.add("d-none");
                    }
                }]
            },
            navigation: {
                itemsEnd: this.generateNavItems(),
                showFilter: false,
                title: Strings.ProjectName,
                onRendering: props => {
                    // Update the navigation properties
                    props.className = "navbar-expand rounded-top";
                    props.type = Components.NavbarTypes.Primary;

                    // Add a logo to the navbar brand
                    let div = document.createElement("div");
                    let text = div.cloneNode() as HTMLDivElement;
                    div.className = "d-flex";
                    text.className = "ms-2";
                    text.append(Strings.ProjectName);
                    div.appendChild(wrenchAdjustableCircle(32, 32, 'brand'));
                    div.appendChild(text);
                    props.brand = div;
                },
                onRendered: (el) => {
                    el.querySelector("a.navbar-brand").classList.add("pe-none");
                }
            },
            subNavigation: {
                onRendering: props => {
                    props.className = "navbar-sub rounded-bottom";
                },
                onRendered: (el) => {
                    el.querySelector("nav.navbar").classList.remove("bg-light");
                },
                itemsEnd: [
                    {
                        text: "Filter Services",
                        onRender: (el, item) => {
                            // Clear the existing button
                            el.innerHTML = "";
                            // Create a span to wrap the icon in
                            let span = document.createElement("span");
                            span.className = "bg-white d-inline-flex ms-2 rounded";
                            el.appendChild(span);

                            // Render a tooltip
                            Components.Tooltip({
                                el: span,
                                content: item.text,
                                placement: Components.TooltipPlacements.Right,
                                btnProps: {
                                    // Render the icon button
                                    className: "p-1 pe-2",
                                    iconClassName: "me-1",
                                    iconType: filterSquare,
                                    iconSize: 24,
                                    isSmall: true,
                                    text: "Filters",
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Show the filter
                                        dashboard.showFilter();
                                    }
                                },
                            });
                        }
                    }
                ]
            },
            tiles: {
                items: DataSource.ListItems,
                bodyField: "ServiceId",
                colSize: Strings.TileColumnSize,
                filterField: "ServiceStatus",
                paginationLimit: Strings.TilePageSize,
                subTitleField: "ServiceStatus",
                titleField: "Title",
                onBodyRendered: (el, item: IListItem) => {
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

                                    // Hide the modal footer
                                    Modal.FooterElement.classList.add("d-none");

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
                },
                onFooterRendered: (el) => {
                    el.classList.add("d-none");
                },
                onHeaderRendered: (el) => {
                    el.classList.add("d-none");
                },
                onPaginationRendered: (el) => {
                    let nav = el.querySelector("nav") as HTMLElement;
                    nav ? nav.classList.remove("pt-2") : null;
                    nav ? nav.classList.add("pt-3") : null;
                }
            }
        });
    }
}