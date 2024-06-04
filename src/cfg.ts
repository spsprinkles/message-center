import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Lists.Main,
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            CustomFields: [
                {
                    name: "Category",
                    title: "Category",
                    type: Helper.SPCfgFieldType.Choice,
                    fillInChoice: true,
                    choices: ["preventOrFixIssue", "PlanForChange", "stayInformed"]
                } as Helper.IFieldInfoChoice,
                {
                    name: "Content",
                    title: "Content",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    unlimited: true
                } as Helper.IFieldInfoNote,
                {
                    name: "IsApproved",
                    title: "Is Approved",
                    type: Helper.SPCfgFieldType.Boolean,
                    defaultValue: "0"
                },
                {
                    name: "IsMajorChange",
                    title: "Is Major Change",
                    type: Helper.SPCfgFieldType.Boolean,
                    defaultValue: "0"
                },
                {
                    name: "Message",
                    title: "Message",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    unlimited: true
                } as Helper.IFieldInfoNote,
                {
                    name: "MessageId",
                    title: "Message ID",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "Platform",
                    title: "Platform",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "PublishedDate",
                    title: "Published Date",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "RoadMapId",
                    title: "Road Map ID",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "Services",
                    title: "Services",
                    type: Helper.SPCfgFieldType.Choice,
                    fillInChoice: true,
                    multi: true,
                    choices: []
                } as Helper.IFieldInfoChoice,
                {
                    name: "Severity",
                    title: "Severity",
                    type: Helper.SPCfgFieldType.Choice,
                    fillInChoice: true,
                    choices: ["normal", "high", "critical"]
                } as Helper.IFieldInfoChoice,
                {
                    name: "Summary",
                    title: "Summary",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    unlimited: true
                } as Helper.IFieldInfoNote,
                {
                    name: "Tags",
                    title: "Tags",
                    type: Helper.SPCfgFieldType.Choice,
                    fillInChoice: true,
                    multi: true,
                    choices: []
                } as Helper.IFieldInfoChoice
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "MessageId", "LinkTitle", "Summary"
                    ]
                }
            ]
        }
    ]
});