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
            TitleFieldDisplayName: "Service",
            CustomFields: [
                {
                    name: "ServiceId",
                    title: "Service ID",
                    type: Helper.SPCfgFieldType.Text,
                    required: true
                },
                {
                    name: "ServiceIssues",
                    title: "Service Issues",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    unlimited: true
                } as Helper.IFieldInfoNote,
                {
                    name: "ServiceStatus",
                    title: "Service Status",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "Operational",
                    required: true,
                    choices: [
                        "Operational", "Investigating", "Degradation",
                        "Interruption", "Restoring service", "Extended recovery",
                        "Investigation suspended", "Restored", "False positive",
                        "Post-incident report published"
                    ]
                } as Helper.IFieldInfoChoice
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "ServiceId", "ServiceStatus"
                    ]
                }
            ]
        }
    ]
});