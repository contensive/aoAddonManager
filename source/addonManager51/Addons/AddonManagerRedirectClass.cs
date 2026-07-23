using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.AddonManager51 {
    /// <summary>
    /// Redirects to the Collection Library addon within the Add-on Manager portal
    /// </summary>
    public class AddonManagerRedirectClass : AddonBaseClass {
        public override object Execute(CPBaseClass cp) {
            try {
                cp.Response.Redirect("?addonguid={A1BCA00C-2657-42A1-8F98-BED1E5A42025}&setPortalGuid={9DF996D7-2280-431E-8D34-2ED059CB988C}");
                return string.Empty;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}
