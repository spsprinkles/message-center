import { ContextInfo, SPTypes } from "gd-sprest-bs";

// App Properties
export interface IAppProps {
    el: HTMLElement;
    context?: any;
    displayMode?: number;
    envType?: number;
    listName?: string;
    moreInfo?: string;
    moreInfoTooltip?: string;
    sortField?: string;
    sourceUrl?: string;
    tileColumnSize?: number;
    tilePageSize?: number;
    timeFormat?: string;
    timeZone?: string;
    title?: string;
}

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (props: IAppProps) => {
    // Set the context
    props.context ? ContextInfo.setPageContext(props.context.pageContext) : null;

    // Update the properties
    Strings.IsClassic = props.envType == SPTypes.EnvironmentType.ClassicSharePoint;
    Strings.SourceUrl = props.sourceUrl || ContextInfo.webServerRelativeUrl;

    // Update the MoreInfo from SPFx title field
    props.moreInfo ? Strings.MoreInfo = props.moreInfo : null;

    // Update the MoreInfo from SPFx title field
    props.moreInfoTooltip ? Strings.MoreInfoTooltip = props.moreInfoTooltip : null;

    // Update the TileColumnSize from SPFx value
    props.tileColumnSize ? Strings.TileColumnSize = props.tileColumnSize : null;

    // Update the TilePageSize from SPFx value, set it to max value if OnlyTiles = true
    props.tilePageSize ? Strings.TilePageSize = props.tilePageSize : null;

    // Update the TimeFormat from SPFx value
    props.timeFormat ? Strings.TimeFormat = props.timeFormat : null;

    // Update the TimeZone from SPFx value
    props.timeZone ? Strings.TimeZone = props.timeZone : null;

    // Update the ProjectName from SPFx title field
    props.title ? Strings.ProjectName = props.title : null;

    // Update the SortField from SPFx title field
    props.sortField ? Strings.SortField = props.sortField : null;
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
    MoreInfo: "More Info",
    MoreInfoTooltip: "Click to view additional details for this item.",
    ProjectName: "Message Center",
    ProjectDescription: "The Message Center app is a solution that reads service messages data from a SharePoint list and presents it to all users with an intuitive interface.",
    SortField: "Modified",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    TileColumnSize: 3,
    TilePageSize: 9,
    TimeFormat: "YYYY-MMM-DD HH:mm:ss zz",
    TimeZone: "America/New_York",
    Version: "0.0.3"
};
export default Strings;