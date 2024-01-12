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
                    name: "Status",
                    title: "Status",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "Draft",
                    required: true,
                    showInNewForm: false,
                    choices: [
                        "Draft", "Submitted", "Rejected", "Pending Approval",
                        "Approved", "Archived"
                    ]
                } as Helper.IFieldInfoChoice
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "Status"
                    ]
                }
            ]
        }
    ]
});