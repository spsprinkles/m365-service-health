import { waitForTheme } from "dattatable";
import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Properties
interface IProps {
    el: HTMLElement;
    context?: any;
    displayMode?: number;
    envType?: number;
    onLoaded?: () => void;
    onlyTiles?: boolean;
    showServices?: string[];
    tileColumnSize?: number;
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
    getServices: () => { return DataSource.Services; },
    onlyTiles: Strings.OnlyTiles,
    render: (props: IProps) => {
        // See if the page context exists
        if (props.context) {
            // Set the context
            setContext(props.context, props.envType, props.sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(props.sourceUrl || ContextInfo.webServerRelativeUrl);
        }

        // Update the OnlyTiles value from SPFx settings
        (typeof (props.onlyTiles) === "undefined") ? null : Strings.OnlyTiles = props.onlyTiles;

        // Update the ShowServices array from SPFx value
        (props.showServices && props.showServices.length > 0) ? Strings.ShowServices = props.showServices : null;

        // Update the TileColumnSize from SPFx value
        props.tileColumnSize ? Strings.TileColumnSize = props.tileColumnSize : null;

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
    tilePageSize: Strings.TilePageSize,
    timeFormat: Strings.TimeFormat,
    timeZone: Strings.TimeZone,
    title: Strings.ProjectName,
    updateTheme: (themeInfo) => {
        // Set the theme
        ThemeManager.setCurrentTheme(themeInfo);
    },
    version: Strings.Version
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