using System;
using Contensive.BaseClasses;

namespace Contensive.AddonManager.Addons {
    /// <summary>
    /// Redirects to the Collection Library addon within the Add-on Manager portal
    /// </summary>
    public class AddonManagerRedirectClass : AddonBaseClass {
        /// <summary>
        /// Redirects the user to the Add-on Manager portal collection library view
        /// </summary>
        public override object Execute(CPBaseClass cp) {
            try {
                cp.Response.Redirect("?addonguid={A1BCA00C-2657-42A1-8F98-BED1E5A42025}&setPortalGuid={9DF996D7-2280-431E-8D34-2ED059CB988C}");
                return string.Empty;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return "<p>An error occurred. Please try again.</p>";
            }
        }
    }
}
