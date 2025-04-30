using System;
using Contensive.BaseClasses;
using Microsoft.VisualBasic.CompilerServices;

namespace Contensive.Addons.AddonManager51 {
    // 
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
        // ====================================================================================================
        // 
        public override object Execute(CPBaseClass CP) {
            try {
                cp = CP;
                return getUpload();
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // ====================================================================================================
        // 
        private string getUpload() {
            string returnResult = "";
            try {
                string Button = cp.Doc.GetText(RequestNameButton);
                if ((Button ?? "") == ButtonCancel) {
                    // 
                    // ----- redirect back to the root
                    cp.Response.Redirect(cp.Site.GetText("adminUrl"));
                } else {
                    // 
                    // -- create form
                    var form = new PortalFramework.FormNameValueRowsClass() {
                        title = "Upload Collection",
                        body = cp.Html.p("Use this form to upload an add-on collection. If the GUID of the add-on matches one already installed on this server, it will be updated. If the GUID is new, it will be added."),
                        description = cp.Html.p("Upload a collection zip file to install the collection on this site. ")
                    };
                    if (!cp.User.IsAdmin) {
                        // 
                        // -- Put up error message
                        form.body += cp.Html.p("You must be an administrator to use this tool.");
                    } else {
                        if (Operators.CompareString(cp.Site.GetText("buildVersion"), cp.Version, false) < 0) {
                            // 
                            // -- database needs to be upgraded
                            form.description += cp.Html.p("WARNING: The site database needs to be upgraded. You should do this before installing addon collections.");
                        }
                        // 
                        // -- Upload tool
                        if ((Button ?? "") == ButtonOK) {
                            // 
                            // -- handle upload
                            if (cp.User.IsDeveloper & cp.Doc.GetBoolean("InstallCore")) {
                                // 
                                // -- Reinstall core collection
                                cp.Content.Delete("Add-on Collections", "ccguid='{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}'");
                                // 
                                // -- use v5 method
                                string ErrorMessage = "";
                                if (!InstallController.installCollectionFromLibrary(cp, "{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}", ref ErrorMessage)) {
                                    form.description += cp.Html.p("ERROR: " + ErrorMessage);
                                }
                            }
                            string uploadFilename = cp.Doc.GetText(constants.rnUploadCollectionFile);
                            if (!string.IsNullOrEmpty(uploadFilename)) {
                                // 
                                // -- version 5.0, separate class so this project can be built with contensive 5.0 reference, but run against contensive 4.1
                                string ErrorMessage = "";
                                bool installDependencies = !cp.Doc.GetBoolean(constants.rnBlockDependencies);
                                if (InstallController.installCollectionFromUpload(cp, constants.rnUploadCollectionFile, ref ErrorMessage)) {
                                    if (!string.IsNullOrEmpty(ErrorMessage)) {
                                        // 
                                        // -- install successful, but a problem
                                        form.body += cp.Html.p("Installation completed with the follow message [" + ErrorMessage + "].");
                                    } else {
                                        // 
                                        // -- install failed
                                        form.body += cp.Html.p("Installation successful.");
                                    }
                                } else {
                                    form.body += cp.Html.p("Error installing collection files, ERROR: " + ErrorMessage);
                                }
                            }
                        }
                        // 
                        // Get Form
                        form.addRow();
                        form.rowName = "Collection Zip File";
                        form.rowValue = cp.Html.InputFile(constants.rnUploadCollectionFile);
                        form.addRow();
                        form.rowName = "Block Dependencies";
                        form.rowValue = cp.Html.CheckBox(constants.rnBlockDependencies);
                        form.addFormHidden("UploadCount", "1");
                        form.addFormButton(ButtonOK);
                        form.addFormButton(ButtonCancel);
                    }
                    returnResult = form.getHtml(cp);
                }
                return returnResult;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}