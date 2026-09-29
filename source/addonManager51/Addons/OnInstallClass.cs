using System;
using Contensive.BaseClasses;

namespace Contensive.AddonManager.Addons {
    /// <summary>
    /// Lifecycle hook that runs when this addon collection is installed, updating the layout template
    /// </summary>
    public class OnInstallClass : AddonBaseClass {
        /// <summary>
        /// Updates the addon manager library layout template on installation
        /// </summary>
        public override object Execute(CPBaseClass CP) {
            try {
                CP.Layout.updateLayout(Constants.guidAddonManagerLibraryListCell, Constants.nameAddonManagerLibraryLisCell, Constants.pathFilenameAddonManagerLibraryLisCell);
                return "";
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                return "";
            }
        }
    }
}