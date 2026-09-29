using System;
using Contensive.AddonManager.Controllers;
using Contensive.BaseClasses;

namespace Contensive.AddonManager.Addons {
    /// <summary>
    /// Handles file upload and installation of addon collections
    /// </summary>
    public class UploadClass : AddonBaseClass {
        //
        // -- injected objects -- do not dispose
        private CPBaseClass cp;
        //
        // -- class scope
        private const string RequestNameButton = "button";
        private const string ButtonCancel = " Cancel ";
        private const string ButtonOK = " OK ";
        //
        /// <summary>
        /// Addon entry point for collection upload
        /// </summary>
        public override object Execute(CPBaseClass CP) {
            try {
                cp = CP;
                return getUpload();
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                return "<p>An error occurred during collection upload. Please try again.</p>";
            }
        }
        //
        private string getUpload() {
            try {
                //
                if (!cp.User.IsAdmin) {
                    return cp.Html.p("You must be an administrator to use this tool.");
                }
                //
                string Button = cp.Doc.GetText(RequestNameButton);
                if ((Button ?? "") == ButtonCancel) {
                    //
                    // ----- redirect back to the root
                    cp.Response.Redirect(cp.Site.GetText("adminUrl"));
                }
                //
                // -- create form
                var form = cp.AdminUI.CreateLayoutBuilderNameValue();
                form.title = "Upload Collection";
                form.body = cp.Html.p("Use this form to upload an add-on collection. If the GUID of the add-on matches one already installed on this server, it will be updated. If the GUID is new, it will be added.");
                form.description = cp.Html.p("Upload a collection zip file to install the collection on this site. ");
                {
                    if (string.Compare(cp.Site.GetText("buildVersion"), cp.Version, StringComparison.Ordinal) < 0) {
                        //
                        // -- database needs to be upgraded
                        form.description += cp.Html.p("WARNING: The site database needs to be upgraded. You should do this before installing addon collections.");
                    }
                    //
                    // -- Upload tool
                    if ((Button ?? "") == ButtonOK) {
                        //
                        // -- handle upload
                        if (cp.User.IsDeveloper && cp.Doc.GetBoolean("InstallCore")) {
                            //
                            // -- Reinstall core collection
                            cp.Content.Delete("Add-on Collections", "ccguid='{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}'");
                            //
                            string ErrorMessage = "";
                            if (!InstallController.installCollectionFromLibrary(cp, "{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}", ref ErrorMessage)) {
                                form.description += cp.Html.p($"ERROR: {ErrorMessage}");
                            }
                        }
                        string uploadFilename = cp.Doc.GetText(Constants.rnUploadCollectionFile);
                        if (!string.IsNullOrEmpty(uploadFilename)) {
                            string ErrorMessage = "";
                            bool installDependencies = !cp.Doc.GetBoolean(Constants.rnBlockDependencies);
                            if (InstallController.installCollectionFromUpload(cp, Constants.rnUploadCollectionFile, ref ErrorMessage)) {
                                if (!string.IsNullOrEmpty(ErrorMessage)) {
                                    //
                                    // -- install successful, but a problem
                                    form.body += cp.Html.p($"Installation completed with the follow message [{ErrorMessage}].");
                                } else {
                                    //
                                    // -- install succeeded
                                    form.body += cp.Html.p("Installation successful.");
                                }
                            } else {
                                form.body += cp.Html.p($"Error installing collection files, ERROR: {ErrorMessage}");
                            }
                        }
                    }
                    //
                    // Get Form
                    form.addRow();
                    form.rowName = "Collection Zip File";
                    form.rowValue = cp.Html.InputFile(Constants.rnUploadCollectionFile);
                    //
                    form.addRow();
                    form.rowName = "Block Dependencies";
                    form.rowValue = cp.Html.CheckBox(Constants.rnBlockDependencies);
                    //
                    form.addFormHidden("UploadCount", "1");
                    form.addFormButton(ButtonOK);
                    form.addFormButton(ButtonCancel);
                }
                return form.getHtml();
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}
