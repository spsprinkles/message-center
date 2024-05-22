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
    AppElementId: "message-center",
    GlobalVariable: "MessageCenter",
    IsClassic: true,
    Lists: {
        Main: "Message Center"
    },
    MaxPageSize: 500,
    MoreInfo: null,
    MoreInfoTooltip: "View more information",
    OnlyTiles: false,
    ProjectName: "Message Center",
    ProjectDescription: "The Message Center app is a solution that reads service messages data from a SharePoint list and presents it to all users with an intuitive interface.",
    ShowServices: null,
    SourceUrl: ContextInfo.webServerRelativeUrl,
    TileColumnSize: 3,
    TileCompact: false,
    TilePageSize: 9,
    TimeFormat: "YYYY-MMM-DD HH:mm:ss zz",
    TimeZone: "America/New_York",
    Version: "0.0.1"
};
export default Strings;