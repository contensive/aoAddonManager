using System;
using Contensive.BaseClasses;
using Microsoft.VisualBasic.CompilerServices;

namespace Contensive.Addons.AddonManager51 {
    // 
    public class OnInstallClass : AddonBaseClass {
        // 
        // ====================================================================================================
        // 
        public override object Execute(CPBaseClass CP) {
            try {
                CP.Layout.updateLayout(constants.guidAddonManagerLibraryListCell, constants.nameAddonManagerLibraryLisCell, constants.pathFilenameAddonManagerLibraryLisCell);
                return "";
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}