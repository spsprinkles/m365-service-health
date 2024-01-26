import { Dashboard, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { checkCircleFill } from "gd-sprest-bs/build/icons/svgs/checkCircleFill";
import { exclamationTriangleFill } from "gd-sprest-bs/build/icons/svgs/exclamationTriangleFill";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { infoCircleFill } from "gd-sprest-bs/build/icons/svgs/infoCircleFill";
import { xCircleFill } from "gd-sprest-bs/build/icons/svgs/xCircleFill";
import * as moment from "moment";
import { DataSource, IListItem } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
import * as common from "./common";
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
                    div.appendChild(common.getIcon(32, 32, 'ServiceHealth'));
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
                bodyField: "ServiceStatus",
                colSize: Strings.TileColumnSize,
                filterField: "ServiceStatus",
                paginationLimit: Strings.TilePageSize,
                showFooter: false,
                subTitleField: "ServiceId",
                titleField: "Title",
                onBodyRendered: (el, item: IListItem) => {
                    let text = el.querySelector("div.card-text") as HTMLDivElement;
                    if (text) {
                        text.classList.add("mt-2");
                        text.innerHTML = "<b>Status:</b> " + common.getStatusTitle(item.ServiceStatus);
                        let p = document.createElement("p");
                        p.className = "mb-0 mt-1 small";
                        p.textContent = common.getStatusDescription(item.ServiceStatus);
                        text.appendChild(p);
                        if (common.getStatusDescription(item.ServiceStatus).length > 75) {
                            Components.Tooltip({
                                content: common.getStatusDescription(item.ServiceStatus),
                                type: Components.TooltipTypes.LightBorder,
                                target: p
                            });
                        }
                    }
                },
                onHeaderRendered: (el, item: IListItem) => {
                    el.classList.add("text-center");
                    let icon = common.getIcon(32, 32, common.getIconName(item.ServiceId));
                    icon.style.pointerEvents = "auto";
                    Components.Tooltip({
                        content: item.Title,
                        type: Components.TooltipTypes.LightBorder,
                        target: icon as any
                    });
                    el.appendChild(icon);
                },
                onPaginationRendered: (el) => {
                    let nav = el.querySelector("nav") as HTMLElement;
                    nav ? nav.classList.remove("pt-2") : null;
                    nav ? nav.classList.add("pt-3") : null;
                },
                onSubTitleRendered: (el, item: IListItem) => {
                    // Clear the element
                    while (el.firstChild) { el.removeChild(el.firstChild); }

                    // Render a div for the status icon
                    let div = document.createElement("div");
                    div.className = "d-inline me-2 mw-fit";
                    el.appendChild(div);

                    // See if issues exist
                    let issues = item.ServiceIssues ? JSON.parse(item.ServiceIssues) : null;
                    if (issues && issues.length > 0) {
                        let advisory = 0;
                        let incident = 0;
                        
                        // Identify if issues are advisory or incidents
                        for (let i = 0; i < issues.length; i++) {
                            if (issues[i].classification == "advisory") { advisory++ };
                            if (issues[i].classification == "incident") { incident++ };
                        }
                        
                        let btnText = "";
                        let tooltip = "View the ";
                        if (incident > 0) {
                            // Render the incident icon
                            div.classList.add("text-warning");
                            div.appendChild(exclamationTriangleFill(20, 20));
                            Components.Tooltip({
                                content: common.getClassificationDescription("incident"),
                                options: { allowHTML: true, maxWidth: 500 },
                                type: Components.TooltipTypes.LightBorder,
                                target: div
                            });
                            if (incident == 1) {
                                btnText = incident + " incident";
                                tooltip += "incident";
                            } else {
                                btnText = incident + " incidents";
                                tooltip += "incidents";
                            }
                            if (advisory > 0) {
                                if (advisory == 1) {
                                    btnText += ", " + advisory + " advisory";
                                    tooltip += " & advisory";
                                } else {
                                    btnText += ", " + advisory + " advisories";
                                    tooltip += " & advisories";
                                }
                            }
                        } else if (advisory > 0) {
                            // Render the advisory icon
                            div.classList.add("text-blue");
                            div.appendChild(infoCircleFill(20, 20));
                            Components.Tooltip({
                                content: common.getClassificationDescription("advisory"),
                                options: { allowHTML: true },
                                type: Components.TooltipTypes.LightBorder,
                                target: div
                            });
                            if (advisory == 1) {
                                btnText = advisory + " advisory";
                                tooltip += "advisory";
                            } else {
                                btnText = advisory + " advisories";
                                tooltip += "advisories";
                            }
                        } else {
                            // Render the error icon
                            div.classList.add("text-danger");
                            div.appendChild(xCircleFill(20, 20));
                            Components.Tooltip({
                                content: common.getClassificationDescription("error"),
                                options: { allowHTML: true },
                                type: Components.TooltipTypes.LightBorder,
                                target: div
                            });
                            if (issues.length == 1) {
                                btnText = issues.length + " error";
                                tooltip += "error";
                            } else {
                                btnText = issues.length + " errors";
                                tooltip += "errors";
                            }
                        }
                        tooltip += " for this service";

                        // Render a button to show the issues
                        Components.Tooltip({
                            el,
                            content: tooltip,
                            btnProps: {
                                className: "p-0",
                                text: btnText,
                                type: Components.ButtonTypes.Link,
                                onClick: () => {
                                    // Clear the modal
                                    Modal.clear();
                                    
                                    // Set the modal header
                                    Modal.setHeader(common.getIcon(28, 28, common.getIconName(item.ServiceId), 'me-2'));
                                    Modal.HeaderElement.append(item.Title + " Issues");
                                    Modal.HeaderElement.classList.add("d-flex");

                                    // Hide the modal footer
                                    Modal.FooterElement.classList.add("d-none");

                                    // Parse the items
                                    let items: Components.IListGroupItem[] = [];
                                    for(let i=0; i<issues.length; i++) {
                                        let issue = issues[i];

                                        // Create the content element
                                        let content = document.createElement("div");

                                        // Set the classification icon
                                        let classIcon = document.createElement("div");
                                        classIcon.className = "d-inline-flex me-1 mw-fit";
                                        if (issue.classification == "incident") {
                                            // Render the incident icon
                                            classIcon.classList.add("text-warning");
                                            classIcon.appendChild(exclamationTriangleFill(20, 20));
                                            Components.Tooltip({
                                                content: common.getClassificationDescription(issue.classification),
                                                options: { allowHTML: true, maxWidth: 500 },
                                                type: Components.TooltipTypes.LightBorder,
                                                target: classIcon
                                            });
                                        } else if (issue.classification == "advisory") {
                                            // Render the advisory icon
                                            classIcon.classList.add("text-blue");
                                            classIcon.appendChild(infoCircleFill(20, 20));
                                            Components.Tooltip({
                                                content: common.getClassificationDescription(issue.classification),
                                                options: { allowHTML: true },
                                                type: Components.TooltipTypes.LightBorder,
                                                target: classIcon
                                            });
                                        } else {
                                            // Render the error icon
                                            classIcon.classList.add("text-danger");
                                            classIcon.appendChild(xCircleFill(20, 20));
                                            Components.Tooltip({
                                                content: common.getClassificationDescription(issue.classification),
                                                options: { allowHTML: true },
                                                type: Components.TooltipTypes.LightBorder,
                                                target: classIcon
                                            });
                                        }
                                        
                                        let issueTitle = document.createElement("h6");
                                        issueTitle.className = "d-inline-flex issue-title mb-0";
                                        issueTitle.innerHTML = common.uppercaseFirst(issue.classification);
                                        
                                        let issueText = document.createElement("p");
                                        issueText.className = "mb-0";
                                        issueText.innerHTML = `
                                            <b>Issue:</b> ${issue.title}<br/>
                                            <b>Details:</b> ${issue.impactDescription}<br/>
                                            <b>Feature:</b> ${issue.feature}<br/>
                                            <b>Group:</b> ${issue.featureGroup}<br/>
                                            <b>Status:</b> ${common.getStatusTitle(issue.status)}<br/>
                                            <b>Start time:</b> ${moment(issue.startDateTime).format(Strings.TimeFormat)}<br/>
                                            <b>Last updated:</b> ${moment(issue.lastModifiedDateTime).format(Strings.TimeFormat)}
                                        `;

                                        // Add the elements to the content
                                        content.appendChild(classIcon);
                                        content.appendChild(issueTitle);
                                        content.appendChild(issueText);

                                        // Add the issue
                                        items.push({ content });
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
                    } else {
                        // Render the healthy icon
                        div.classList.add("text-success");
                        div.appendChild(checkCircleFill(20, 20));
                        Components.Tooltip({
                            content: common.getClassificationDescription("healthy"),
                            options: { allowHTML: true },
                            type: Components.TooltipTypes.LightBorder,
                            target: div
                        });
                        el.append("Healthy");
                        el.classList.add("fw-normal");
                    }
                }
            }
        });
    }
}