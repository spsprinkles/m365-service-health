import { waitForTheme } from "dattatable";
import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { getIcon, invertIconColor } from "./common";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import { Services } from "./propertyPane";
import { Security } from "./security";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Properties
interface IProps {
    el: HTMLElement;
    context?: any;
    darkTheme?: boolean;
    displayMode?: number;
    envType?: number;
    listName?: string;
    moreInfo?: string;
    moreInfoTooltip?: string;
    onLoaded?: () => void;
    onlyTiles?: boolean;
    showServices?: string[];
    tileColumnSize?: number;
    tileCompact?: boolean;
    tilePageSize?: number;
    timeFormat?: string;
    timeZone?: string;
    title?: string;
    sourceUrl?: string;
}

// Create the global variable for this solution
const GlobalVariable = {
    App: null,
    Configuration,
    description: Strings.ProjectDescription,
    getLogo: () => { return invertIconColor(getIcon(28, 28, 'ServiceHealth', 'logo me-2')); },
    getServices: () => { return DataSource.Services; },
    listName: Strings.Lists.Main,
    moreInfoTooltip: Strings.MoreInfoTooltip,
    propertyPaneServices: Services,
    onlyTiles: Strings.OnlyTiles,
    render: (props: IProps) => {
        // See if the page context exists
        if (props.context) {
            // Set the context
            setContext(props.context, props.envType, props.sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(props.sourceUrl || ContextInfo.webServerRelativeUrl);

            // See if the list name is set
            if (props.listName) {
                // Update the configuration
                Strings.Lists.Main = props.listName;
                Configuration._configuration.ListCfg[0].ListInformation.Title = props.listName;
            }
        }

        // Update the dark theme flag
        Strings.IsDarkTheme = props.darkTheme ? true : false;

        // Update the MoreInfo from SPFx title field
        props.moreInfo ? Strings.MoreInfo = props.moreInfo : Strings.MoreInfo = null;

        // Update the MoreInfo from SPFx title field
        props.moreInfoTooltip ? Strings.MoreInfoTooltip = props.moreInfoTooltip : null;

        // Update the OnlyTiles value from SPFx settings
        (typeof (props.onlyTiles) === "undefined") ? null : Strings.OnlyTiles = props.onlyTiles;

        // Update the ShowServices array from SPFx value
        (props.showServices && props.showServices.length > 0) ? Strings.ShowServices = props.showServices : Strings.ShowServices = null;

        // Update the TileColumnSize from SPFx value
        props.tileColumnSize ? Strings.TileColumnSize = props.tileColumnSize : null;

        // Update the TileCompact value from SPFx settings
        (typeof (props.tileCompact) === "undefined") ? null : Strings.TileCompact = props.tileCompact;

        // Update the TilePageSize from SPFx value, set it to max value if OnlyTiles = true
        props.tilePageSize ? (Strings.OnlyTiles ? Strings.TilePageSize = Strings.MaxPageSize : Strings.TilePageSize = props.tilePageSize) : null;

        // Update the TimeFormat from SPFx value
        props.timeFormat ? Strings.TimeFormat = props.timeFormat : null;

        // Update the TimeZone from SPFx value
        props.timeZone ? Strings.TimeZone = props.timeZone : null;

        // Update the ProjectName from SPFx title field
        props.title ? Strings.ProjectName = props.title : null;

        // Initialize the application
        DataSource.init().then(
            // Success
            () => {
                // Wait for the theme to be loaded
                waitForTheme().then(() => {
                    // Create the application
                    GlobalVariable.App = new App(props.el);

                    // Call the load event
                    props.onLoaded ? props.onLoaded() : null;
                });
            },

            // Error
            () => {
                // See if the user has the correct permissions
                Security.hasPermissions().then(hasPermissions => {
                    // See if the user has permissions
                    if (hasPermissions) {
                        // Show the installation modal
                        InstallationModal.show();
                    }
                });
            }
        );
    },
    tileColumnSize: Strings.TileColumnSize,
    tileCompact: Strings.TileCompact,
    tilePageSize: Strings.TilePageSize,
    timeFormat: Strings.TimeFormat,
    timeZone: Strings.TimeZone,
    title: Strings.ProjectName,
    updateTheme: (themeInfo) => {
        // Set the theme
        ThemeManager.setCurrentTheme(themeInfo);
    }
};

// Make is available in the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Remove the extra border spacing on the webpart in classic mode
    let contentBox = document.querySelector("#contentBox table.ms-core-tableNoSpace");
    contentBox ? contentBox.classList.remove("ms-webpartPage-root") : null;

    // Render the application
    GlobalVariable.render({ el: elApp });
}