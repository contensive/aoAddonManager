using System;
using System.Data;
using Contensive.AddonManager.Controllers;
using Contensive.BaseClasses;
using Contensive.Models.Db;

namespace Contensive.AddonManager.Addons {
    /// <summary>
    /// uninstall a collection
    /// </summary>
    public class UninstallClass : AddonBaseClass {
        // 
        // injected objects -- do not dispose
        // 
        private CPBaseClass cp;
        // 
        // class scope
        // 
        private const string RequestNameButton = "button";
        // 
        private const string ButtonCancel = " Cancel ";
        private const string ButtonOK = " OK ";
        //
        // =====================================================================================
        /// <summary>
        /// addon api
        /// </summary>
        /// <param name="CP"></param>
        /// <returns></returns>
        public override object Execute(CPBaseClass CP) {
            try {
                cp = CP;
                return getUnistall();
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                return "<p>An error occurred during collection uninstall. Please try again.</p>";
            }
        }
        // 
        // ==========================================================================================================================================
        /// <summary>
        /// Addon Manager,  This is a form that lets you upload an addon
        /// Eventually, this should be substituted with a "Addon Manager Addon" - so the interface can be improved with Contensive recompile
        /// </summary>
        /// <returns></returns>
        private string getUnistall() {
            string returnResult = "";
            try {
                var form = cp.AdminUI.CreateLayoutBuilderList("{3092AD1E-9660-4104-81D3-C29210C3183C}");
                form.callbackAddonGuid = "{3092AD1E-9660-4104-81D3-C29210C3183C}";
                form.paginationPageSizeDefault = 50;
                //
                if (!cp.User.IsAdmin) {
                    return cp.Html.p("You must be an administrator to use this tool.");
                }
                //
                string SiteKey = cp.Site.GetText("sitekey", "");
                if (string.IsNullOrEmpty(SiteKey)) {
                    SiteKey = cp.Utils.CreateGuid();
                    cp.Site.SetProperty("sitekey", SiteKey);
                }
                //
                bool DbUpToDate = string.Compare(cp.Site.GetText("buildVersion"), cp.Version, StringComparison.Ordinal) >= 0;
                string Button = cp.Doc.GetText(RequestNameButton);
                if ((Button ?? "") == ButtonCancel) {
                    //
                    // ----- redirect back to the root
                    //
                    cp.Response.Redirect(cp.Site.GetText("adminUrl"));
                } else {
                    {
                        // installFolder = "CollectionUpload" & cp.Utils.CreateGuid().Replace("{", "").Replace("-", "").Replace("}", "")
                        // InstallPath = cp.Site.PhysicalFilePath & installFolder & "\"
                        if ((Button ?? "") == ButtonOK) {
                            // 
                            // ---------------------------------------------------------------------------------------------
                            // Delete collections
                            // Before deleting each addon, make sure it is not in another collection
                            // ---------------------------------------------------------------------------------------------
                            // 
                            int cnt = cp.Doc.GetInteger("accnt");
                            if (cnt > 0) {
                                var loopTo = cnt - 1;
                                int Ptr;
                                for (Ptr = 0; Ptr <= loopTo; Ptr++) {
                                    if (cp.Doc.GetBoolean($"ac{Ptr}")) {
                                        int TargetCollectionID = cp.Doc.GetInteger($"acID{Ptr}");
                                        InstallController.uninstallCollection(cp, TargetCollectionID);
                                    }
                                }
                            }
                        }
                        // 
                        // --------------------------------------------------------------------------------
                        // Current Collections Tab
                        // --------------------------------------------------------------------------------
                        // 
                        form.columnCaption = "Del";
                        form.columnCaptionClass = $"{Constants.AfwStyles.afwTextAlignCenter} {Constants.AfwStyles.afwWidth50px}";
                        form.columnCellClass = Constants.AfwStyles.afwTextAlignCenter;
                        form.columnDownloadable = false;
                        form.columnName = "";
                        form.columnSortable = false;
                        form.columnVisible = true;
                        form.addColumn();
                        //
                        form.columnCaption = "Name";
                        form.columnCaptionClass = Constants.AfwStyles.afwTextAlignLeft;
                        form.columnCellClass = Constants.AfwStyles.afwTextAlignLeft;
                        form.columnDownloadable = false;
                        form.columnName = "name";
                        form.columnSortable = true;
                        form.columnVisible = true;
                        form.addColumn();
                        //
                        // Build SQL criteria with search filter
                        //
                        string sqlCriteria = "";
                        if (!string.IsNullOrEmpty(form.sqlSearchTerm)) {
                            sqlCriteria = $"(name like {cp.Db.EncodeSQLTextLike(form.sqlSearchTerm)})";
                        }
                        //
                        // Build SQL order by with sort
                        //
                        string sqlOrderBy = "name";
                        if (!string.IsNullOrEmpty(form.sortField)) {
                            sqlOrderBy = form.sortField;
                            if (form.sortDirection == "desc") {
                                sqlOrderBy += " desc";
                            }
                        }
                        //
                        // Get total record count for pagination
                        //
                        string searchWhere = string.IsNullOrEmpty(sqlCriteria) ? "" : $" and ({sqlCriteria})";
                        string sqlCount = $"select count(*) from ccaddoncollections where (1=1){searchWhere}";
                        using (DataTable dt = cp.Db.ExecuteQuery(sqlCount)) {
                            if (dt?.Rows != null && dt.Rows.Count == 1) {
                                form.recordCount = cp.Utils.EncodeInteger(dt.Rows[0][0]);
                            }
                        }
                        //
                        // Get paginated data
                        //
                        var addonCollectionList = DbBaseModel.createList<AddonCollectionModel>(cp, sqlCriteria, sqlOrderBy, form.paginationPageSize, form.paginationPageNumber);
                        int rowPtr = 0;
                        foreach (var item in addonCollectionList) {
                            form.addRow();
                            form.setCell($"{cp.Html.CheckBox($"AC{rowPtr}")}{cp.Html.Hidden($"ACID{rowPtr}", item.id.ToString())}");
                            form.setCell(item.name);
                            rowPtr += 1;
                        }
                        form.addFormButton(ButtonCancel);
                        form.addFormButton(ButtonOK);
                        form.addFormHidden("accnt", rowPtr.ToString());
                    }
                    // 
                    // Output the Add-on
                    // 
                    form.title = "Uninstall Collections";
                    form.description = "To remove collections, select them from the list and click the Uninstall button.";
                    if (!DbUpToDate) {
                        form.description += "<div style=\"Margin-left:50px\">Warning: The site's Database needs to be upgraded.</div>";
                    }
                    returnResult = form.getHtml();
                }
                return returnResult;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }

    }
}