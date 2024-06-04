import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { clipboard } from "gd-sprest-bs/build/icons/svgs/clipboard";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
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