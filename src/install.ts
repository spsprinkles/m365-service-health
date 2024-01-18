import { InstallationRequired } from "dattatable";
import { Components } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { Security } from "./security";
import Strings from "./strings";

/**
 * Installation Modal
 */
export class InstallationModal {
    // Shows the modal
    static show(showFl: boolean = false) {
        // See if an installation is required
        InstallationRequired.requiresInstall({ cfg: Configuration }).then(installFl => {
            // See if an install is required
            if (installFl || showFl) {
                // Show the dialog
                InstallationRequired.showDialog({
                    onFooterRendered: el => {
                        // See if the security group doesn't exist
                        if (showFl) {
                            // Add the custom install button
                            Components.Tooltip({
                                el,
                                content: "Manage the security groups",
                                type: Components.ButtonTypes.OutlinePrimary,
                                btnProps: {
                                    text: "Security",
                                    isDisabled: !InstallationRequired.ListsExist,
                                    onClick: () => {
                                        // Show the security dialog
                                        Security.show(() => {
                                            // Refresh the page
                                            window.location.reload();
                                        });
                                    }
                                }
                            });
                        }
                    }
                });
            } else {
                // Log
                console.error("[" + Strings.ProjectName + "] Error initializing the solution.");
            }
        });
    }
}