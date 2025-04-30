using System;
using Contensive.Addons.PortalFramework;
using Contensive.BaseClasses;
using Contensive.Models.Db;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Contensive.Addons.AddonManager51 {
    public class ExportClass : AddonBaseClass {
        // 
        // ====================================================================================================
        /// <summary>
        /// export collection
        /// </summary>
        /// <param name="CP"></param>
        /// <returns></returns>
        public override object Execute(CPBaseClass CP) {
            try {
                string Button = CP.Doc.GetText(constants.RequestNameButton);
                if ((Button ?? "") == constants.ButtonCancel) {
                    // 
                    // ----- redirect back to the root
                    CP.Response.Redirect(CP.Site.GetText("adminUrl"));
                    return string.Empty;
                }
                // 
                // -- create form
                var form = new FormNameValueRowsClass() {
                    title = "Export Collection",
                    body = CP.Html.p("Use this tool to create an Add-on Collection zip file that can be used to install a collection on another site.")
                };
                if (!CP.User.IsAdmin) {
                    // 
                    // -- Put up error message
                    form.body += CP.Html.p("You must be an administrator to use this tool.");
                    return form.getHtml(CP);
                }
                if (Operators.CompareString(CP.Site.GetText("buildVersion"), CP.Version, false) < 0) {
                    // 
                    // -- database needs to be upgraded
                    form.description += CP.Html.p("Warning: The site's Database needs to be upgraded. You should do this before installing addons.");
                }
                // 
                // -- Upload tool
                if ((Button ?? "") == constants.ButtonOK) {
                    // 
                    // -- export
                    int CollectionID = CP.Doc.GetInteger(constants.RequestNameCollectionID);
                    var addonCollection = DbBaseModel.create<AddonCollectionModel>(CP, CollectionID);
                    if (addonCollection is null) {
                        // 
                        // -- collection not found
                        CP.UserError.Add("The collection file you selected could not be found. Please select another.");
                    } else {
                        // 
                        // -- build collection zip file and return file
                        string CollectionFilename = "";
                        string userError = "";
                        // 
                        // -- attempt new method
                        CP.Addon.ExportCollection(CollectionID, ref CollectionFilename, ref userError);
                        if (!string.IsNullOrEmpty(userError)) {
                            // 
                            // -- errors during export
                            form.body = CP.Html.div(CP.Html.p("ERRORS during export: ") + CP.Html.ul(userError));
                        } else if (!CP.UserError.OK()) {
                            // 
                            // -- errors during export
                            form.body = CP.Html.div(CP.Html.p("ERRORS during export: ") + CP.Html.ul(CP.UserError.GetList()));
                        } else {
                            // 
                            // -- success
                            form.body += CP.Html.p("Export Successful");
                            form.body += CP.Html.p("Click <a href=\"" + CP.Http.CdnFilePathPrefixAbsolute + Strings.Replace(CollectionFilename, @"\", "/") + "\">here</a> to download the collection file.</p>");
                        }
                    }
                }
                // 
                // Get Form
                form.addRow();
                form.rowName = "Add-on Collection";
                form.rowValue = CP.Html.SelectContent(constants.RequestNameCollectionID, "0", "Add-on Collections", "", "Select Collection To Export", "form-control");
                form.addFormHidden("UploadCount", "1");
                form.addFormButton(constants.ButtonOK);
                form.addFormButton(constants.ButtonCancel);
                return form.getHtml(CP);
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}