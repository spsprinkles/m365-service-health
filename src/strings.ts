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
        Main: "M365 Service Health"
    },
    MaxPageSize: 500,
    OnlyTiles: false,
    ProjectName: "M365 Service Health",
    ProjectDescription: "A webpart showing M365 Service Health for all users.",
    ShowServices: null,
    SourceUrl: ContextInfo.webServerRelativeUrl,
    TileColumnSize: 3,
    TileCompact: false,
    TilePageSize: 9,
    TimeFormat: "YYYY-MMM-DD HH:mm:ss zz",
    TimeZone: "America/New_York",
    Version: "0.0.3"
};
export default Strings;