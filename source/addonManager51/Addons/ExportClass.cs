using System;
using Contensive.BaseClasses;
using Contensive.Models.Db;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Contensive.Addons.AddonManager51 {
    /// <summary>
    /// Export Addon Collection
    /// </summary>
    public class ExportClass : AddonBaseClass {
        // 
        // ====================================================================================================
        /// <summary>
        /// export collection
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        public override object Execute(CPBaseClass cp) {
            try {
                string Button = cp.Doc.GetText(constants.RequestNameButton);
                if ((Button ?? "") == constants.ButtonCancel) {
                    // 
                    // ----- redirect back to the root
                    cp.Response.Redirect(cp.Site.GetText("adminUrl"));
                    return string.Empty;
                }
                // 
                // -- create form
                var form = cp.AdminUI.CreateLayoutBuilderNameValue();
                form.title = "Export Collection";
                form.body = cp.Html.p("Use this tool to create an Add-on Collection zip file that can be used to install a collection on another site.");
                if (!cp.User.IsAdmin) {
                    // 
                    // -- Put up error message
                    form.body += cp.Html.p("You must be an administrator to use this tool.");
                    return form.getHtml();
                }
                if (Operators.CompareString(cp.Site.GetText("buildVersion"), cp.Version, false) < 0) {
                    // 
                    // -- database needs to be upgraded
                    form.description += cp.Html.p("Warning: The site's Database needs to be upgraded. You should do this before installing addons.");
                }
                // 
                // -- Upload tool
                if ((Button ?? "") == constants.ButtonOK) {
                    // 
                    // -- export
                    int CollectionID = cp.Doc.GetInteger(constants.RequestNameCollectionID);
                    var addonCollection = DbBaseModel.create<AddonCollectionModel>(cp, CollectionID);
                    if (addonCollection is null) {
                        // 
                        // -- collection not found
                        cp.UserError.Add("The collection file you selected could not be found. Please select another.");
                    } else {
                        // 
                        // -- build collection zip file and return file
                        string CollectionFilename = "";
                        string userError = "";
                        // 
                        // -- attempt new method
                        cp.Addon.ExportCollection(CollectionID, ref CollectionFilename, ref userError);
                        if (!string.IsNullOrEmpty(userError)) {
                            // 
                            // -- errors during export
                            form.body = cp.Html.div(cp.Html.p("ERRORS during export: ") + cp.Html.ul(userError));
                        } else if (!cp.UserError.OK()) {
                            // 
                            // -- errors during export
                            form.body = cp.Html.div(cp.Html.p("ERRORS during export: ") + cp.Html.ul(cp.UserError.GetList()));
                        } else {
                            // 
                            // -- success
                            form.body += cp.Html.p("Export Successful");
                            form.body += cp.Html.p("Click <a href=\"" + cp.Http.CdnFilePathPrefixAbsolute + Strings.Replace(CollectionFilename, @"\", "/") + "\">here</a> to download the collection file.</p>");
                        }
                    }
                }
                // 
                // Get Form
                form.addRow();
                form.rowName = "Add-on Collection";
                form.rowValue = cp.Html.SelectContent(constants.RequestNameCollectionID, "0", "Add-on Collections", "", "Select Collection To Export", "form-control");
                form.addFormHidden("UploadCount", "1");
                form.addFormButton(constants.ButtonOK);
                form.addFormButton(constants.ButtonCancel);
                return form.getHtml();
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}