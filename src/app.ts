import { Dashboard, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { checkCircleFill } from "gd-sprest-bs/build/icons/svgs/checkCircleFill";
import { exclamationTriangleFill } from "gd-sprest-bs/build/icons/svgs/exclamationTriangleFill";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { infoCircleFill } from "gd-sprest-bs/build/icons/svgs/infoCircleFill";
import { xCircleFill } from "gd-sprest-bs/build/icons/svgs/xCircleFill";
import { DataSource, IListItem } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
import { getIcon } from "./common";
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
                    div.appendChild(getIcon(32, 32, 'ServiceHealth'));
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
                        text.innerHTML = "<b>Status:</b> " + item.ServiceStatus;
                    }
                },
                onHeaderRendered: (el, item: IListItem) => {
                    el.classList.add("text-center");
                    let serviceId = item.ServiceId;
                    switch (item.ServiceId) {
                        case "cloudappsecurity":
                            serviceId = "SecurityCenter";
                            break;
                        case "DynamicsCRM":
                            serviceId = "Dynamics";
                            break;
                        case "Lync":
                            serviceId = "Skype";
                            break;
                        case "Microsoft365Defender":
                            serviceId = "Defender";
                            break;
                        case "MicrosoftFlow":
                        case "MicrosoftFlowM365":
                            serviceId = "PowerAutomate";
                            break;
                        case "microsoftteams":
                            serviceId = "Teams";
                            break;
                        case "MobileDeviceManagement":
                            serviceId = "MDM";
                            break;
                        case "O365Client":
                            serviceId = "Office";
                            break;
                        case "officeonline":
                            serviceId = "OfficeOnline";
                            break;
                        case "OneDriveForBusiness":
                            serviceId = "OneDrive";
                            break;
                        case "OrgLiveID":
                            serviceId = "Entra";
                            break;
                        case "OSDPPlatform":
                            serviceId = "M365";
                            break;
                        case "PAM":
                            serviceId = "KeyVault";
                            break;
                        case "PowerAppsM365":
                            serviceId = "PowerApps";
                            break;
                        case "PowerBIcom":
                            serviceId = "PowerBI";
                            break;
                        case "ProjectForTheWeb":
                            serviceId = "Project";
                            break;
                        case "RMS":
                            serviceId = "AIP";
                            break;
                        case "SwayEnterprise":
                            serviceId = "Sway";
                            break;
                    }
                    el.appendChild(getIcon(32, 32, serviceId));
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
                                content: "<b>Incident:</b> A critical service issue with noticeable user impact.",
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
                            div.classList.add("text-primary");
                            div.appendChild(infoCircleFill(20, 20));
                            Components.Tooltip({
                                content: "<b>Advisory:</b> A minor service issue with limited impact.",
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
                                content: "<b>Error:</b> A major service issue with a broad user impact.",
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
                    } else {
                        // Render the healthy icon
                        div.classList.add("text-success");
                        div.appendChild(checkCircleFill(20, 20));
                        Components.Tooltip({
                            content: "<b>Healthy:</b> No incidents or advisories.",
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