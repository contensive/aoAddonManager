using System;
using Contensive.Addons.PortalFramework;
using Contensive.BaseClasses;
using Contensive.Models.Db;
using Microsoft.VisualBasic.CompilerServices;

namespace Contensive.Addons.AddonManager51 {
    // 
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
        private struct NavigatorType {
            public string Name;
            public string NameSpacex;
        }

        private struct Collection2Type {
            public int AddonCnt;
            public string[] AddonGuid;
            public string[] AddonName;
            public int MenuCnt;
            public string[] Menus;
            public int NavigatorCnt;
            public NavigatorType[] Navigators;
        }
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
                throw;
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

                bool DisplaySystem;
                bool DbUpToDate;
                string GuidFieldName;
                int AddonNavigatorID;
                string TargetCollectionName;
                int addonid;
                int cnt;
                int Ptr;
                string BodyHTML;
                var cs = cp.CSNew();
                string Button;
                string ButtonList;
                string CollectionFilename = "";
                var Doc = new System.Xml.XmlDocument();
                string status = "";
                bool collectionsToBeInstalledFromFolder;
                int TargetCollectionID;
                string SiteKey;
                var form = new ReportListClass(cp);
                // 
                SiteKey = cp.Site.GetText("sitekey", "");
                if (string.IsNullOrEmpty(SiteKey)) {
                    SiteKey = cp.Utils.CreateGuid();
                    cp.Site.SetProperty("sitekey", SiteKey);
                }
                // 
                DbUpToDate = Operators.CompareString(cp.Site.GetText("buildVersion"), cp.Version, false) >= 0;
                Button = cp.Doc.GetText(RequestNameButton);
                collectionsToBeInstalledFromFolder = false;
                GuidFieldName = "ccguid";
                if ((Button ?? "") == ButtonCancel) {
                    // 
                    // ----- redirect back to the root
                    // 
                    cp.Response.Redirect(cp.Site.GetText("adminUrl"));
                } else {
                    if (!cp.User.IsAdmin) {
                        // 
                        // ----- Put up error message
                        // 
                        ButtonList = ButtonCancel;
                        BodyHTML = cp.Html.p("You must be an administrator to use this tool.");
                    } else {
                        // installFolder = "CollectionUpload" & cp.Utils.CreateGuid().Replace("{", "").Replace("-", "").Replace("}", "")
                        // InstallPath = cp.Site.PhysicalFilePath & installFolder & "\"
                        if ((Button ?? "") == ButtonOK) {
                            // 
                            // ---------------------------------------------------------------------------------------------
                            // Delete collections
                            // Before deleting each addon, make sure it is not in another collection
                            // ---------------------------------------------------------------------------------------------
                            // 
                            cnt = cp.Doc.GetInteger("accnt");
                            if (cnt > 0) {
                                var loopTo = cnt - 1;
                                for (Ptr = 0; Ptr <= loopTo; Ptr++) {
                                    if (cp.Doc.GetBoolean("ac" + Ptr)) {
                                        TargetCollectionID = cp.Doc.GetInteger("acID" + Ptr);
                                        TargetCollectionName = cp.Content.GetRecordName("Add-on Collections", TargetCollectionID);
                                        // 
                                        // Clean up rules associating this collection to other objects
                                        // 
                                        cp.Content.Delete("Add-on Collection CDef Rules", "collectionid=" + TargetCollectionID);
                                        cp.Content.Delete("Add-on Collection Module Rules", "collectionid=" + TargetCollectionID);
                                        // 
                                        // Delete any addons from this collection
                                        // 
                                        if (cs.Open("add-ons", "collectionid=" + TargetCollectionID)) {
                                            do {
                                                // 
                                                // Clean up the rules that might have pointed to the addon
                                                // 
                                                addonid = cs.GetInteger("id");
                                                cp.Content.Delete("Admin Menuing", "addonid=" + addonid);
                                                cp.Content.Delete("Shared Styles Add-on Rules", "addonid=" + addonid);
                                                cp.Content.Delete("Add-on Scripting Module Rules", "addonid=" + addonid);
                                                cp.Content.Delete("Add-on Include Rules", "addonid=" + addonid);
                                                cp.Content.Delete("Add-on Include Rules", "includedaddonid=" + addonid);
                                                cs.GoNext();
                                            }
                                            while (cs.OK());
                                        }
                                        cs.Close();
                                        cp.Content.Delete("add-ons", "collectionid=" + TargetCollectionID);
                                        // 
                                        // Delete the navigator entry for the collection under 'Add-ons'
                                        // 
                                        if (TargetCollectionID > 0) {
                                            AddonNavigatorID = 0;
                                            cs.Open("Navigator Entries", "name='Manage Add-ons' and ((parentid=0)or(parentid is null))");
                                            if (cs.OK()) {
                                                AddonNavigatorID = cs.GetInteger("ID");
                                            }
                                            cs.Close();
                                            if (AddonNavigatorID > 0) {
                                                GenericController.getForm_AddonManager_DeleteNavigatorBranch(cp, TargetCollectionName, AddonNavigatorID);
                                            }
                                            // 
                                            // Now delete the Collection record
                                            // 
                                            cp.Content.Delete("Add-on Collections", "id=" + TargetCollectionID);
                                            // 
                                            // Delete Navigator Entries set as installed by the collection (this may be all that is needed)
                                            // 
                                            cp.Content.Delete("Navigator Entries", "installedbycollectionid=" + TargetCollectionID);
                                        }
                                    }
                                }
                            }
                        }
                        // 
                        // --------------------------------------------------------------------------------
                        // Current Collections Tab
                        // --------------------------------------------------------------------------------
                        // 
                        form.addColumn();
                        form.columnCaption = "Del";
                        form.columnCaptionClass = AfwStyles.afwTextAlignCenter + " " + AfwStyles.afwWidth50px;
                        form.columnCellClass = AfwStyles.afwTextAlignCenter;
                        form.columnDownloadable = false;
                        form.columnName = "";
                        form.columnSortable = false;
                        form.columnVisible = true;
                        // 
                        form.addColumn();
                        form.columnCaption = "Name";
                        form.columnCaptionClass = AfwStyles.afwTextAlignLeft;
                        form.columnCellClass = AfwStyles.afwTextAlignLeft;
                        form.columnDownloadable = false;
                        form.columnName = "";
                        form.columnSortable = false;
                        form.columnVisible = true;
                        // 
                        DisplaySystem = false;
                        var addonCollectionList = DbBaseModel.createList<AddonCollectionModel>(cp, "", "name");
                        int rowPtr = 0;
                        foreach (var item in addonCollectionList) {
                            form.addRow();
                            form.setCell(cp.Html.CheckBox("AC" + rowPtr) + cp.Html.Hidden("ACID" + rowPtr, item.id.ToString()));
                            form.setCell(item.name);
                            rowPtr += 1;
                        }
                        form.addFormButton(ButtonCancel);
                        form.addFormButton(ButtonOK);
                        form.addFormHidden("accnt", addonCollectionList.Count.ToString());
                    }
                    // 
                    // Output the Add-on
                    // 
                    form.name = "Uninstall Collections";
                    form.description = "To remove collections, select them from the list and click the Uninstall button.";
                    if (!DbUpToDate) {
                        form.description += "<div style=\"Margin-left:50px\">Warning: The site's Database needs to be upgraded.</div>";
                    }
                    if (!string.IsNullOrEmpty(status)) {
                        form.description += "<div style=\"Margin-left:50px\">" + status + "</div>";
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