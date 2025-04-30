using System;
using Contensive.Addons.PortalFramework;
using Contensive.BaseClasses;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Contensive.Addons.AddonManager51 {
    // 
    public class CollectionLibraryClass : AddonBaseClass {
        // 
        // injected objects -- do not dispose
        // 
        private CPBaseClass cp;
        // 
        // class scope
        // 
        private const string RequestNameButton = "button";
        private const string CollectionListRootNode = "collectionlist";
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
        private readonly int CollectionCnt;
        private readonly Collection2Type[] Collections;
        private const string guidAddonManagerLibraryListCell = "{9767F464-3728-4B7D-904B-3442D7FD03BE}";
        private const string guidAddonManagerActiveX = "{1DC06F61-1837-419B-AF36-D5CC41E1C9FD}";
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
                return getCollectionLibrary();
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // ==========================================================================================================================================
        /// <summary>
        /// Addon Manager, This is a form that lets you upload an addon, Eventually, this should be substituted with a "Addon Manager Addon" - so the interface can be improved with Contensive recompile
        /// </summary>
        /// <returns></returns>
        private string getCollectionLibrary() {
            string returnResult = "";
            try {
                var form = new LayoutBuilderSimple();
                bool showAddon;
                bool DbUpToDate;
                string GuidFieldName;
                string ErrorMessage = "";
                bool UpgradeOK;
                System.Xml.XmlDocument LibCollections;
                string installFolder;
                string[] LibGuids;
                string CollectionGUID = "";
                string CollectionVersion;
                string CollectionContensiveVersion = "";
                int cnt;
                int Ptr;
                var RowPtr = default(int);
                int PageNumber;
                int ColumnCnt;
                string[] ColCaption;
                string[] ColAlign;
                string[] ColWidth;
                bool[] ColSortable;
                string PreTableCopy;
                string BodyHTML = "";
                string UserError;
                string Button;
                string ButtonList;
                string CollectionFilename = "";
                var Doc = new System.Xml.XmlDocument();
                string status = "";
                bool collectionsToBeInstalledFromFolder;
                string InstallLibCollectionList = "";
                string InstallPath;
                string SiteKey;
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
                    return "";
                }
                if (!cp.User.IsAdmin) {
                    // 
                    // ----- Put up error message
                    // 
                    ButtonList = ButtonCancel;
                    BodyHTML = cp.Html.p("You must be an administrator to use this tool.");
                } else {
                    // 
                    PreTableCopy = "Use this form to upload an add-on collection. If the GUID of the add-on matches one already installed on this server, it will be updated. If the GUID is new, it will be added.";
                    installFolder = "CollectionUpload" + cp.Utils.CreateGuid().Replace("{", "").Replace("-", "").Replace("}", "");
                    InstallPath = installFolder + @"\";
                    if ((Button ?? "") == ButtonOK) {
                        // 
                        // ---------------------------------------------------------------------------------------------
                        // Download and install Collections from the Collection Library
                        // ---------------------------------------------------------------------------------------------
                        // 
                        InstallLibCollectionList = "";
                        if (!string.IsNullOrEmpty(cp.Doc.GetText("LibraryRow"))) {
                            Ptr = cp.Doc.GetInteger("LibraryRow");
                            CollectionGUID = cp.Doc.GetText("LibraryRowguid" + Ptr);
                            InstallLibCollectionList = InstallLibCollectionList + "," + CollectionGUID;
                        }
                        // 
                        if (!string.IsNullOrEmpty(InstallLibCollectionList)) {
                            InstallLibCollectionList = Strings.Mid(InstallLibCollectionList, 2);
                            LibGuids = Strings.Split(InstallLibCollectionList, ",");
                            cnt = Information.UBound(LibGuids) + 1;
                            var loopTo = cnt - 1;
                            for (Ptr = 0; Ptr <= loopTo; Ptr++) {
                                // 
                                // -- use v5 method
                                UpgradeOK = InstallController.installCollectionFromLibrary(cp, LibGuids[Ptr], ref ErrorMessage);
                                if (!UpgradeOK) {
                                    form.failMessage += ErrorMessage;
                                    continue;
                                }
                                if (!string.IsNullOrEmpty(ErrorMessage)) {
                                    form.warningMessage += ErrorMessage;
                                    continue;
                                }
                            }
                        }
                    }
                    // 
                    // --------------------------------------------------------------------------------
                    // Get Form
                    // --------------------------------------------------------------------------------
                    // Get the Collection Library tab
                    // --------------------------------------------------------------------------------
                    // 
                    ColumnCnt = 4;
                    PageNumber = 1;
                    ColCaption = new string[4];
                    ColAlign = new string[4];
                    ColWidth = new string[4];
                    ColSortable = new bool[4];
                    // 
                    ColCaption[0] = "Install";
                    ColAlign[0] = "center";
                    ColWidth[0] = "50";
                    ColSortable[0] = false;
                    // 
                    ColCaption[1] = "Name";
                    ColAlign[1] = "left";
                    ColWidth[1] = "200";
                    ColSortable[1] = false;
                    // 
                    ColCaption[2] = "Last&nbsp;Updated";
                    ColAlign[2] = "right";
                    ColWidth[2] = "200";
                    ColSortable[2] = false;
                    // 
                    ColCaption[3] = "Description";
                    ColAlign[3] = "left";
                    ColWidth[3] = "99%";
                    ColSortable[3] = false;
                    // 
                    LibCollections = new System.Xml.XmlDocument();
                    LibCollections.Load("http://support.contensive.com/GetCollectionList?iv=" + cp.Version + "&key=" + cp.Utils.EncodeRequestVariable(SiteKey) + "&name=" + cp.Utils.EncodeRequestVariable(cp.Site.Name) + "&primaryDomain=" + cp.Utils.EncodeRequestVariable(cp.Site.DomainPrimary));
                    if (true) {
                        if ((LibCollections.DocumentElement.Name.ToLower() ?? "") != (CollectionListRootNode.ToLower() ?? "")) {
                            UserError = "There was an error reading the Collection Library file. The '" + CollectionListRootNode + "' element was not found.";
                            status = status + "<BR>" + UserError;
                            cp.UserError.Add(UserError);
                        } else {
                            // 
                            // Go through file to validate the XML, and build status message -- since service process can not communicate to user
                            // 
                            RowPtr = 0;
                            // Content = ""
                            string cellTemplate = My.Resources.Resources.AddonManagerLibraryListCell;
                            foreach (System.Xml.XmlNode CDef_Node in LibCollections.DocumentElement.ChildNodes) {
                                string Cell = cellTemplate;
                                string CollectionImageLink = "";
                                string CollectionCheckbox = "";
                                string CollectionName = "";
                                var CollectionModifiedDate = DateTime.MinValue;
                                string CollectionModifiedDateCaption = "";
                                string CollectionDescription = "";
                                string CollectionHelpLink = "";
                                string CollectionDemoLink = "";
                                switch (Strings.LCase(CDef_Node.Name) ?? "") {
                                    case "collection": {
                                            // 
                                            // Read the collection
                                            // 
                                            foreach (System.Xml.XmlNode CollectionNode in CDef_Node.ChildNodes) {
                                                switch (CollectionNode.Name.ToLower() ?? "") {
                                                    case "name": {
                                                            // 
                                                            // Name
                                                            // 
                                                            CollectionName = CollectionNode.InnerText;
                                                            break;
                                                        }
                                                    case "helplink": {
                                                            // 
                                                            // helpLink
                                                            // 
                                                            CollectionHelpLink = CollectionNode.InnerText;
                                                            break;
                                                        }
                                                    case "demolink": {
                                                            // 
                                                            // demoLink
                                                            // 
                                                            CollectionDemoLink = CollectionNode.InnerText;
                                                            break;
                                                        }
                                                    case "guid": {
                                                            // 
                                                            // Guid
                                                            // 
                                                            CollectionGUID = CollectionNode.InnerText;
                                                            break;
                                                        }
                                                    case "version": {
                                                            // 
                                                            // Version
                                                            // 
                                                            CollectionVersion = CollectionNode.InnerText;
                                                            break;
                                                        }
                                                    case "description": {
                                                            // 
                                                            // Version
                                                            // 
                                                            CollectionDescription = CollectionNode.InnerText;
                                                            break;
                                                        }
                                                    case "imagelink": {
                                                            // 
                                                            // Version
                                                            // 
                                                            CollectionImageLink = CollectionNode.InnerText;
                                                            break;
                                                        }
                                                    case "contensiveversion": {
                                                            // 
                                                            // Version
                                                            // 
                                                            CollectionContensiveVersion = CollectionNode.InnerText;
                                                            break;
                                                        }
                                                    case "lastchangedate": {
                                                            break;
                                                        }
                                                    // 
                                                    // last change - legacy field used by 4.1 to auto install, no longer updated
                                                    // 
                                                    case "lastmodifieddate": {
                                                            // 
                                                            // last modified
                                                            // 
                                                            CollectionModifiedDate = DateTime.MinValue;
                                                            if (Information.IsDate(CollectionNode.InnerText)) {
                                                                CollectionModifiedDate = Conversions.ToDate(CollectionNode.InnerText);
                                                            }
                                                            if (CollectionModifiedDate <= DateTime.MinValue) {
                                                                CollectionModifiedDateCaption = "unknown";
                                                            } else {
                                                                CollectionModifiedDateCaption = CollectionModifiedDate.Date.ToShortDateString();
                                                            }

                                                            break;
                                                        }
                                                }
                                            }
                                            if (string.IsNullOrEmpty(CollectionImageLink)) {
                                                CollectionImageLink = "/addonManager/libraryNoImage.jpg";
                                            }
                                            if (string.IsNullOrEmpty(CollectionModifiedDateCaption)) {
                                                CollectionModifiedDateCaption = "unknown";
                                            }
                                            if (string.IsNullOrEmpty(CollectionDescription)) {
                                                CollectionDescription = "No description is available for this add-on collection.";
                                            }
                                            showAddon = false;
                                            if (!string.IsNullOrEmpty(CollectionName)) {
                                                if (!string.IsNullOrEmpty(CollectionGUID)) {
                                                    var cs = cp.CSNew();
                                                    var modifiedDate = DateTime.MinValue;
                                                    bool IsOnSite = false;
                                                    if (cs.Open("Add-on Collections", GuidFieldName + "=" + cp.Db.EncodeSQLText(CollectionGUID), "", true, "ID,ModifiedDate")) {
                                                        modifiedDate = cs.GetDate("ModifiedDate");
                                                        IsOnSite = true;
                                                    }
                                                    cs.Close();
                                                    CollectionCheckbox = "<input TYPE=\"CheckBox\" NAME=\"LibraryRow\" VALUE=\"" + RowPtr + "\" onClick=\"clearLibraryRows('" + RowPtr + "');\">" + cp.Html.Hidden("LibraryRowGuid" + RowPtr, CollectionGUID) + cp.Html.Hidden("LibraryRowName" + RowPtr, CollectionName);
                                                    if (!IsOnSite) {
                                                        // 
                                                        // -- not installed
                                                        showAddon = true;
                                                        CollectionCheckbox += "&nbsp;Install";
                                                    } else if (modifiedDate >= CollectionModifiedDate) {
                                                        // 
                                                        // -- up to date, reinstall
                                                        showAddon = true;
                                                        CollectionCheckbox += "&nbsp;Reinstall";
                                                    } else {
                                                        // 
                                                        // -- old version, upgrade
                                                        showAddon = true;
                                                        CollectionCheckbox += "&nbsp;Upgrade";
                                                    }
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(CollectionDemoLink)) {
                                                CollectionDescription = CollectionDescription + "<div class=\"amDemoLink\"><a target=\"_blank\" href=\"" + CollectionDemoLink + "\">Demo</a></div>";
                                            }
                                            if (!string.IsNullOrEmpty(CollectionHelpLink)) {
                                                CollectionDescription = CollectionDescription + "<div class=\"amHelpLink\"><a target=\"_blank\" href=\"" + CollectionHelpLink + "\">Reference</a></div>";
                                            }
                                            if (showAddon) {
                                                Cell = Strings.Replace(Cell, "##imageLink##", CollectionImageLink);
                                                Cell = Strings.Replace(Cell, "##checkbox##", CollectionCheckbox);
                                                Cell = Strings.Replace(Cell, "##name##", CollectionName);
                                                Cell = Strings.Replace(Cell, "##date##", CollectionModifiedDateCaption);
                                                Cell = Strings.Replace(Cell, "##description##", CollectionDescription);
                                                BodyHTML += Cell;
                                            }
                                            RowPtr += 1;
                                            break;
                                        }
                                }
                            }
                        }
                        BodyHTML = "" + constants.cr + "<script language=\"JavaScript\">" + "function clearLibraryRows(r) {" + "var c,p;" + "c=document.getElementsByName('LibraryRow');" + "for (p=0;p<c.length;p++){" + "if(c[p].value!=r)c[p].checked=false;" + "}" + "" + "}" + "</script>" + "<input type=hidden name=LibraryCnt value=\"" + RowPtr + "\">" + constants.cr + "<div style=\"width:100%\">" + BodyHTML + constants.cr + "</div>" + "";














                        form.body = BodyHTML;
                        form.description = "Select Add-ons to install from the Contensive Add-on Library. Please select only one at a time. Click OK to install the selected Add-on. The site may need to be stopped during the installation, but will be available again in approximately one minute";

                        // BodyHTML = AdminUI.GetEditPanel(main, True, "Add-on Collection Library", "Select Add-ons to install from the Contensive Add-on Library. Please select only one at a time. Click OK to install the selected Add-on. The site may need to be stopped during the installation, but will be available again in approximately one minute.", BodyHTML)
                        // BodyHTML = BodyHTML & cp.Html.Hidden("AOCnt", RowPtr)
                        // Call main.AddLiveTabEntry("<NOBR>Collection&nbsp;Library</NOBR>", BodyHTML, "ccAdminTab")
                    }
                    // 
                    // --------------------------------------------------------------------------------
                    // Build Page from tabs
                    // --------------------------------------------------------------------------------
                    // 
                    form.addFormButton(ButtonCancel);
                    form.addFormButton(ButtonOK);
                }
                // 
                // Output the Add-on
                // 
                form.title = "Add-on Library";
                form.description = "";
                if (!DbUpToDate) {
                    form.description += "<div style=\"Margin-left:50px\">The Add-on Manager is disabled because this site's Database needs to be upgraded.</div>";
                }
                if (!string.IsNullOrEmpty(status)) {
                    form.description = form.description + "<div style=\"Margin-left:50px\">" + status + "</div>";
                }
                return form.getHtml(cp);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return returnResult;
        }
    }
}