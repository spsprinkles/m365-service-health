import { Dashboard, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
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

    // Returns the M365 icon as an SVG element
    private getM365Icon(height?, width?, className?) {
        // Set the default values
        if (height === void 0) { height = 32; }
        if (width === void 0) { width = 32; }

        // Get the icon element
        let elDiv = document.createElement("div");
        elDiv.innerHTML = "<svg viewBox='0 0 48 48' fill='none' xmlns='http://www.w3.org/2000/svg'><path d='M20.0842 3.02588L19.8595 3.16179C19.5021 3.37799 19.1654 3.61972 18.8512 3.88385L19.4993 3.42798H25L26 11L21 16L16 19.4754V23.4829C16 26.2819 17.4629 28.8774 19.8574 30.3268L25.1211 33.5129L14 40.0002H11.8551L7.85737 37.5804C5.46286 36.131 4 33.5355 4 30.7365V17.2606C4 14.4607 5.46379 11.8645 7.85952 10.4154L19.8595 3.15687C19.9339 3.11189 20.0088 3.06823 20.0842 3.02588Z' fill='url(#part0)'></path><path d='M20.0842 3.02588L19.8595 3.16179C19.5021 3.37799 19.1654 3.61972 18.8512 3.88385L19.4993 3.42798H25L26 11L21 16L16 19.4754V23.4829C16 26.2819 17.4629 28.8774 19.8574 30.3268L25.1211 33.5129L14 40.0002H11.8551L7.85737 37.5804C5.46286 36.131 4 33.5355 4 30.7365V17.2606C4 14.4607 5.46379 11.8645 7.85952 10.4154L19.8595 3.15687C19.9339 3.11189 20.0088 3.06823 20.0842 3.02588Z' fill='url(#part1)'></path><path d='M32 19V23.4803C32 26.2793 30.5371 28.8748 28.1426 30.3242L16.1426 37.5878C13.6878 39.0737 10.6335 39.1273 8.1355 37.7487L19.8573 44.844C22.4039 46.3855 25.5959 46.3855 28.1426 44.844L40.1426 37.5803C42.5371 36.1309 43.9999 33.5354 43.9999 30.7364V27.5L42.9999 26L32 19Z' fill='url(#part2)'></path><path d='M32 19V23.4803C32 26.2793 30.5371 28.8748 28.1426 30.3242L16.1426 37.5878C13.6878 39.0737 10.6335 39.1273 8.1355 37.7487L19.8573 44.844C22.4039 46.3855 25.5959 46.3855 28.1426 44.844L40.1426 37.5803C42.5371 36.1309 43.9999 33.5354 43.9999 30.7364V27.5L42.9999 26L32 19Z' fill='url(#part3)'></path><path d='M40.1405 10.4153L28.1405 3.15678C25.6738 1.66471 22.6021 1.61849 20.0979 3.01811L19.8595 3.16231C17.4638 4.61143 16 7.20757 16 10.0075V19.4914L19.8595 17.1568C22.4051 15.6171 25.5949 15.6171 28.1405 17.1568L40.1405 24.4153C42.4613 25.8192 43.9076 28.2994 43.9957 30.9985C43.9986 30.9113 44 30.824 44 30.7364V17.2605C44 14.4606 42.5362 11.8644 40.1405 10.4153Z' fill='url(#part4)'></path><path d='M40.1405 10.4153L28.1405 3.15678C25.6738 1.66471 22.6021 1.61849 20.0979 3.01811L19.8595 3.16231C17.4638 4.61143 16 7.20757 16 10.0075V19.4914L19.8595 17.1568C22.4051 15.6171 25.5949 15.6171 28.1405 17.1568L40.1405 24.4153C42.4613 25.8192 43.9076 28.2994 43.9957 30.9985C43.9986 30.9113 44 30.824 44 30.7364V17.2605C44 14.4606 42.5362 11.8644 40.1405 10.4153Z' fill='url(#part5)'></path><path d='M4.00428 30.9984C4.00428 30.9984 4.00428 30.9984 4.00428 30.9984Z' fill='url(#part6)'></path><path d='M4.00428 30.9984C4.00428 30.9984 4.00428 30.9984 4.00428 30.9984Z' fill='url(#part7)'></path><defs><radialGradient id='part0' cx='0' cy='0' r='1' gradientUnits='userSpaceOnUse' gradientTransform='translate(17.4186 10.6383) rotate(110.528) scale(33.3657 58.1966)'><stop offset='0.06441' stop-color='#AE7FE2'></stop><stop offset='1' stop-color='#0078D4'></stop></radialGradient><linearGradient id='part1' x1='17.5119' y1='37.8685' x2='12.7513' y2='29.6347' gradientUnits='userSpaceOnUse'><stop stop-color='#114A8B'></stop><stop offset='1' stop-color='#0078D4' stop-opacity='0'></stop></linearGradient><radialGradient id='part2' cx='0' cy='0' r='1' gradientUnits='userSpaceOnUse' gradientTransform='translate(10.4299 36.3511) rotate(-8.36717) scale(31.0503 20.5108)'><stop offset='0.133928' stop-color='#D59DFF'></stop><stop offset='1' stop-color='#5E438F'></stop></radialGradient><linearGradient id='part3' x1='40.3566' y1='25.3768' x2='35.2552' y2='32.6916' gradientUnits='userSpaceOnUse'><stop stop-color='#493474'></stop><stop offset='1' stop-color='#8C66BA' stop-opacity='0'></stop></linearGradient><radialGradient id='part4' cx='0' cy='0' r='1' gradientUnits='userSpaceOnUse' gradientTransform='translate(41.0552 26.504) rotate(-165.772) scale(24.9228 41.9552)'><stop offset='0.0584996' stop-color='#50E6FF'></stop><stop offset='1' stop-color='#436DCD'></stop></radialGradient><linearGradient id='part5' x1='16.9758' y1='3.05655' x2='24.4868' y2='3.05655' gradientUnits='userSpaceOnUse'><stop stop-color='#2D3F80'></stop><stop offset='1' stop-color='#436DCD' stop-opacity='0'></stop></linearGradient><radialGradient id='part6' cx='0' cy='0' r='1' gradientUnits='userSpaceOnUse' gradientTransform='translate(41.0552 26.504) rotate(-165.772) scale(24.9228 41.9552)'><stop offset='0.0584996' stop-color='#50E6FF'></stop><stop offset='1' stop-color='#436DCD'></stop></radialGradient><linearGradient id='part7' x1='16.9758' y1='3.05655' x2='24.4868' y2='3.05655' gradientUnits='userSpaceOnUse'><stop stop-color='#2D3F80'></stop><stop offset='1' stop-color='#436DCD' stop-opacity='0'></stop></linearGradient></defs></svg>";
        let icon = elDiv.firstChild as SVGImageElement;
        if (icon) {
            // See if a class name exists
            if (className) {
                // Parse the class names
                let classNames = className.split(' ');
                for (var i = 0; i < classNames.length; i++) {
                    // Add the class name
                    icon.classList.add(classNames[i]);
                }
            } else {
                icon.classList.add("icon-svg");
            }

            // Set the height/width
            height ? icon.setAttribute("height", (height).toString()) : null;
            width ? icon.setAttribute("width", (width).toString()) : null;

            // Hide the icon as non-interactive content from the accessibility API
            icon.setAttribute("aria-hidden", "true");

            // Update the styling
            icon.style.pointerEvents = "none";

            // Support for IE
            icon.setAttribute("focusable", "false");
        }

        // Return the icon
        return icon;
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
                    div.appendChild(this.getM365Icon(32, 32, "m365-logo"));
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