import { ContextInfo, SPTypes } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, envType?: number, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the properties
    Strings.IsClassic = envType == SPTypes.EnvironmentType.ClassicSharePoint;
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "m365-service-health",
    GlobalVariable: "M365ServiceHealth",
    IsClassic: true,
    Lists: {
        Main: "M365ServiceHealth"
    },
    ProjectName: "M365 Service Health",
    ProjectDescription: "A webpart showing M365 Service Health for all users.",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.1"
};
export default Strings;