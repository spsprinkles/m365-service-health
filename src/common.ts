import { CustomIconTypes, CustomIcons } from "gd-sprest-bs";

const classificationDescription = new Map<string, string>([
    ["advisory", "<b>Advisory:</b> A minor service issue with limited impact."],
    ["error", "<b>Error:</b> A major service issue with a broad user impact."],
    ["healthy", "<b>Healthy:</b> No incidents or advisories."],
    ["incident", "<b>Incident:</b> A critical service issue with noticeable user impact."]
]);

const statusDescription = new Map<string, string>([
    ["serviceOperational", "The service is healthy and no issues have been identified."],
    ["investigating", "A potential issue was identified and more information is being gathered about what's going on and the scope of impact."],
    ["restoringService", "The cause of the issue has been identified, and action is being taken to bring the service back to a healthy state."],
    ["verifyingService", "The action has been taken to mitigate the issue and we have verified that the service is healthy."],
    ["serviceRestored", "The corrective action has resolved the underlying problem and the service has been restored to a healthy state. To find out what went wrong, view the issue details."],
    ["postIncidentReviewPublished", "A post-incident report for a specific issue that includes root cause information has been published, with next steps to ensure a similar issue doesn't reoccur."],
    ["serviceDegradation", "An issue is confirmed that may affect use of a service or feature. You might see this status if a service is performing more slowly than usual, there are intermittent interruptions, or if a feature isn't working, for example."],
    ["serviceInterruption", "You'll see this status if an issue is determined to affect the ability for users to access the service. In this case, the issue is significant and can be reproduced consistently."],
    ["extendedRecovery", "This status indicates that corrective action is in progress to restore the service to most users but will take some time to reach all the affected systems. You might also see this status if a temporary fix is made to reduce impact while a permanent fix is waiting to be applied."],
    ["falsePositive", "After a detailed investigation, the service is confirmed to be healthy and operating as designed. No impact to the service was observed or the cause of the incident originated outside of the service. Incidents and advisories with this status appear in the history view until they expire (after the period of time stated in the final post for that event)."],
    ["investigationSuspended", "If our detailed investigation of a potential issue results in a request for additional information from customers to allow the service team to investigate further, you'll see this status. If service team needs you to act, they'll let you know what data or logs they need."]
]);

const statusTitle = new Map<string, string>([
    ["serviceOperational", "Operational"],
    ["investigating", "Investigating"],
    ["restoringService", "Restoring Service"],
    ["verifyingService", "Verifying Service"],
    ["serviceRestored", "Service Restored"],
    ["postIncidentReviewPublished", "Post-incident Report Published"],
    ["serviceDegradation", "Service Degradation"],
    ["serviceInterruption", "Service Interruption"],
    ["extendedRecovery", "Extended Recovery"],
    ["falsePositive", "False Positive"],
    ["investigationSuspended", "Investigation Suspended"]
]);

// Returns a description of the classification value
export function getClassificationDescription(classification?: string) {
    // Return the classification description
    return (classification && classificationDescription.has(classification)) ? classificationDescription.get(classification) : classificationDescription.get('error');
}

// Returns an icon as an SVG element
export function getIcon(height?, width?, iconName?, className?): SVGImageElement {
    // Get the icon element
    let elIcon: HTMLElement = null;
    switch (iconName) {
        case "AIP":
            elIcon = CustomIcons(CustomIconTypes.aIP, height, width, className);
            break;
        case "Bookings":
            elIcon = CustomIcons(CustomIconTypes.bookings, height, width, className);
            break;
        case "Dataverse":
            elIcon = CustomIcons(CustomIconTypes.dataverse, height, width, className);
            break;
        case "Defender":
            elIcon = CustomIcons(CustomIconTypes.defender, height, width, className);
            break;
        case "Dynamics":
            elIcon = CustomIcons(CustomIconTypes.dynamics, height, width, className);
            break;
        case "Entra":
            elIcon = CustomIcons(CustomIconTypes.entra, height, width, className);
            break;
        case "Exchange":
            elIcon = CustomIcons(CustomIconTypes.exchange, height, width, className);
            break;
        case "Forms":
            elIcon = CustomIcons(CustomIconTypes.forms, height, width, className);
            break;
        case "Intune":
            elIcon = CustomIcons(CustomIconTypes.intune, height, width, className);
            break;
        case "KeyVault":
            elIcon = CustomIcons(CustomIconTypes.keyVault, height, width, className);
            break;
        case "M365":
            elIcon = CustomIcons(CustomIconTypes.m365, height, width, className);
            break;
        case "MDM":
            elIcon = CustomIcons(CustomIconTypes.mDM, height, width, className);
            break;
        case "Office":
            elIcon = CustomIcons(CustomIconTypes.office, height, width, className);
            break;
        case "OfficeOnline":
            elIcon = CustomIcons(CustomIconTypes.officeOnline, height, width, className);
            break;
        case "OneDrive":
            elIcon = CustomIcons(CustomIconTypes.oneDrive, height, width, className);
            break;
        case "Planner":
            elIcon = CustomIcons(CustomIconTypes.planner, height, width, className);
            break;
        case "PowerApps":
            elIcon = CustomIcons(CustomIconTypes.powerApps, height, width, className);
            break;
        case "PowerAutomate":
            elIcon = CustomIcons(CustomIconTypes.powerAutomate, height, width, className);
            break;
        case "PowerBI":
            elIcon = CustomIcons(CustomIconTypes.powerBI, height, width, className);
            break;
        case "PowerPlatform":
            elIcon = CustomIcons(CustomIconTypes.powerPlatform, height, width, className);
            break;
        case "Project":
            elIcon = CustomIcons(CustomIconTypes.project, height, width, className);
            break;
        case "SecurityCenter":
            elIcon = CustomIcons(CustomIconTypes.securityCenter, height, width, className);
            break;
        case "ServiceHealth":
            elIcon = CustomIcons(CustomIconTypes.serviceHealth, height, width, className);
            break;
        case "SharePoint":
            elIcon = CustomIcons(CustomIconTypes.sharePoint, height, width, className);
            break;
        case "Skype":
            elIcon = CustomIcons(CustomIconTypes.skype, height, width, className);
            break;
        case "Stream":
            elIcon = CustomIcons(CustomIconTypes.stream, height, width, className);
            break;
        case "Sway":
            elIcon = CustomIcons(CustomIconTypes.sway, height, width, className);
            break;
        case "Teams":
            elIcon = CustomIcons(CustomIconTypes.teams, height, width, className);
            break;
        case "Viva":
            elIcon = CustomIcons(CustomIconTypes.viva, height, width, className);
            break;
        default:
            elIcon = CustomIcons(CustomIconTypes.m365, height, width, className);
            break;
    }

    // Hide the icon as non-interactive content from the accessibility API
    elIcon.setAttribute("aria-hidden", "true");

    // Update the styling
    elIcon.style.pointerEvents = "none";

    // Support for IE
    elIcon.setAttribute("focusable", "false");

    // Return the icon
    return elIcon as any;
}

// Creates a mapping from the serviceId to the proper icon name
export function getIconName(serviceId: string) {
    let svcId = serviceId;
    switch (serviceId) {
        case "cloudappsecurity":
            svcId = "SecurityCenter";
            break;
        case "DynamicsCRM":
            svcId = "Dynamics";
            break;
        case "Lync":
            svcId = "Skype";
            break;
        case "Microsoft365Defender":
            svcId = "Defender";
            break;
        case "MicrosoftDataverse":
            svcId = "Dataverse";
            break;
        case "MicrosoftFlow":
        case "MicrosoftFlowM365":
            svcId = "PowerAutomate";
            break;
        case "microsoftteams":
            svcId = "Teams";
            break;
        case "MobileDeviceManagement":
            svcId = "MDM";
            break;
        case "O365Client":
            svcId = "Office";
            break;
        case "officeonline":
            svcId = "OfficeOnline";
            break;
        case "OneDriveForBusiness":
            svcId = "OneDrive";
            break;
        case "OrgLiveID":
            svcId = "Entra";
            break;
        case "OSDPPlatform":
            svcId = "M365";
            break;
        case "PAM":
            svcId = "KeyVault";
            break;
        case "PowerAppsM365":
            svcId = "PowerApps";
            break;
        case "PowerBIcom":
            svcId = "PowerBI";
            break;
        case "ProjectForTheWeb":
        case "ProjectOnline":
            svcId = "Project";
            break;
        case "RMS":
            svcId = "AIP";
            break;
        case "SwayEnterprise":
            svcId = "Sway";
            break;
    }
    return svcId;
}

// Returns a description of the status value
export function getStatusDescription(statusValue?) {
    // Return the status description
    return (statusValue && statusDescription.has(statusValue)) ? statusDescription.get(statusValue) : statusDescription.get('serviceOperational');
}

// Returns the title of the status value
export function getStatusTitle(statusValue?) {
    // Return the status title
    return (statusValue && statusTitle.has(statusValue)) ? statusTitle.get(statusValue) : statusTitle.get('serviceOperational');
}

// Returns the ServiceHealth icon with the colors inverted
export function invertIconColor(icon: SVGImageElement): SVGImageElement {
    // Invert the theme colors in the icon
    icon.innerHTML = icon.innerHTML.replace('fill="var(--sp-primary-button-text, #ffffff)"', 'fill="var(--sp-theme-primary, #0078d4)"').replace('fill="var(--sp-theme-primary, #0078d4)" style="stroke-width:1.02321"', 'fill="var(--sp-primary-button-text, #ffffff)" style="stroke-width:1.02321"');
    return icon;
}

// Returns a boolean indicating if the urlString is valid
export function isValidUrl(urlString) {
    var urlPattern = new RegExp('^(https?:\\/\\/)?' + // validate protocol
        '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|' + // validate domain name
        '((\\d{1,3}\\.){3}\\d{1,3}))' + // validate OR ip (v4) address
        '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*' + // validate port and path
        '(\\?[;&a-z\\d%_.~+=-]*)?' + // validate query string
        '(\\#[-a-z\\d_]*)?$', 'i'); // validate fragment locator
    return !!urlPattern.test(urlString);
}

// Returns a word with the first letter capitalized
export function uppercaseFirst(word: string) {
    return word.charAt(0).toUpperCase() + word.slice(1);
}