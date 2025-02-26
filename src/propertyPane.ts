import { PropertyPane } from "gd-sprest-bs";
import { DataSource } from "./ds";

export const Services = (key: string, label: string, properties) => {
    // Compute the dropdown items for the property pane
    let items = [];
    DataSource.Services.forEach((value, key) => {
        items.push({
            text: value,
            value: key
        })
    });

    // Return the multi dropdown checkbox
    return PropertyPane.MultiDropdownCheckbox(key, {
        label,
        description: "This is a custom control that doesn't affect this app.",
        placeholder: "Select Services",
        properties,
        tooltip: "This is my tooltip",
        items
    });
}