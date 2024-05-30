import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { clipboard } from "gd-sprest-bs/build/icons/svgs/clipboard";
import { InstallationRequired } from "dattatable";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import Strings, { IAppProps, setContext } from "./strings";

// Styling
import "./styles.scss";


// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    getLogo: () => { return clipboard(28, 28); },
    render: (props: IAppProps) => {
        // Set the context
        setContext(props);

        // Update the configuration
        Configuration.setWebUrl(props.sourceUrl || ContextInfo.webServerRelativeUrl);
        props.listName ? Configuration._configuration.ListCfg[0].ListInformation.Title = props.listName : null;

        // Initialize the application
        DataSource.init().then(
            // Success
            () => {
                // Load the current theme and apply it to the components
                ThemeManager.load(true).then(() => {
                    // Create the application
                    new App(props.el);
                });
            },

            // Error
            () => {
                // See if an installation is required
                InstallationRequired.requiresInstall({ cfg: Configuration }).then(installFl => {
                    // See if an install is required
                    if (installFl) {
                        // Show the dialog
                        InstallationRequired.showDialog();
                    } else {
                        // Log
                        console.error("[" + Strings.ProjectName + "] Error initializing the solution.");
                    }
                });
            }
        );
    },
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
    // Render the application
    GlobalVariable.render({ el: elApp });
}