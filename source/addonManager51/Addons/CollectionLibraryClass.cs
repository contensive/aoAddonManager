using Contensive.BaseClasses;
using Contensive.Models.Db;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using System;
using System.Collections.Generic;
using System.Net;
using System.Runtime.CompilerServices;

namespace Contensive.Addons.AddonManager51 {
    /// <summary>
    /// This is the Add-on Library. 
    /// </summary>
    public class CollectionLibraryClass : AddonBaseClass {
        //
        private CPBaseClass cp;
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
            try {
                var form = cp.AdminUI.CreateLayoutBuilder();
                // 
                string SiteKey = cp.Site.GetText("sitekey", "");
                if (string.IsNullOrEmpty(SiteKey)) {
                    SiteKey = cp.Utils.CreateGuid();
                    cp.Site.SetProperty("sitekey", SiteKey);
                }
                // 
                string Button = cp.Doc.GetText(_Constants.RequestNameButton);
                string ButtonRemove = cp.Doc.GetText(_Constants.ButtonRemove);
                if ((Button ?? "") == _Constants.ButtonCancel) {
                    // 
                    // ----- redirect back to the root
                    // 
                    cp.Response.Redirect(cp.Site.GetText("adminUrl"));
                    return "";
                }
                string status = "";
                if (!cp.User.IsAdmin) {
                    return cp.Html.p("You must be an administrator to use this tool.");
                } else {
                    string installFolder = "CollectionUpload" + cp.Utils.CreateGuid().Replace("{", "").Replace("-", "").Replace("}", "");
                    string InstallPath = installFolder + @"\";
                    string ErrorMessage = "";
                    bool UpgradeOK;
                    string CollectionGUID = "";
                    if (!string.IsNullOrEmpty(ButtonRemove ?? "")) {
                        //
                        // -- uninstall
                        var collection = DbBaseModel.create<AddonCollectionModel>(cp, ButtonRemove);
                        InstallController.UninstallCollection(cp, collection.id);
                    } else if ((Button ?? "") == _Constants.ButtonOK) {
                        //
                        // -- legacy
                        // 
                        // ---------------------------------------------------------------------------------------------
                        // Download and install Collections from the Collection Library
                        // ---------------------------------------------------------------------------------------------
                        // 
                        string InstallLibCollectionList = "";
                        int Ptr;
                        if (!string.IsNullOrEmpty(cp.Doc.GetText("LibraryRow"))) {
                            Ptr = cp.Doc.GetInteger("LibraryRow");
                            CollectionGUID = cp.Doc.GetText("LibraryRowguid" + Ptr);
                            InstallLibCollectionList = InstallLibCollectionList + "," + CollectionGUID;
                        }
                        // 
                        if (!string.IsNullOrEmpty(InstallLibCollectionList)) {
                            InstallLibCollectionList = Strings.Mid(InstallLibCollectionList, 2);
                            string[] LibGuids = Strings.Split(InstallLibCollectionList, ",");
                            int cnt = Information.UBound(LibGuids) + 1;
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
                    } else if (!string.IsNullOrEmpty(Button ?? "")) {
                        //
                        // -- button value is the guid of the collection to install
                        // 
                        // -- use v5 method
                        UpgradeOK = InstallController.installCollectionFromLibrary(cp, Button, ref ErrorMessage);
                        if (!UpgradeOK) {
                            form.failMessage += ErrorMessage;
                        } else if (!string.IsNullOrEmpty(ErrorMessage)) {
                            form.warningMessage += ErrorMessage;
                        }
                    }


                    // 
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    System.Xml.XmlDocument LibCollections = new System.Xml.XmlDocument();
                    LibCollections.Load("https://support.contensive.com/GetCollectionList?iv=" + cp.Version + "&key=" + cp.Utils.EncodeRequestVariable(SiteKey) + "&name=" + cp.Utils.EncodeRequestVariable(cp.Site.Name) + "&primaryDomain=" + cp.Utils.EncodeRequestVariable(cp.Site.DomainPrimary));
                    //
                    LibraryViewModel viewModel = new LibraryViewModel();
                    if ((LibCollections.DocumentElement.Name.ToLower() ?? "") != (_Constants.CollectionListRootNode.ToLower() ?? "")) {
                        string UserError = "There was an error reading the Collection Library file. The '" + _Constants.CollectionListRootNode + "' element was not found.";
                        status = status + "<BR>" + UserError;
                        cp.UserError.Add(UserError);
                    } else {
                        var RowPtr = default(int);
                        // 
                        // Go through file to validate the XML, and build status message -- since service process can not communicate to user
                        // 
                        RowPtr = 0;
                        // Content = ""
                        string cellTemplate = My.Resources.Resources.AddonManagerLibraryListCell;

                        foreach (System.Xml.XmlNode CDef_Node in LibCollections.DocumentElement.ChildNodes) {
                            string CollectionImageLink = "";
                            string CollectionCheckbox = "";
                            string CollectionName = "";
                            var CollectionModifiedDate = DateTime.MinValue;
                            string CollectionModifiedDateCaption = "";
                            string CollectionDescription = "";
                            string CollectionHelpLink = "";
                            string CollectionDemoLink = "";
                            string CollectionContensiveVersion = "";
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
                                                        string CollectionVersion = CollectionNode.InnerText;
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
                                                        // Minimum Contensive version required
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
                                        bool showAddon = false;
                                        bool isInstall = false;
                                        bool isUpgrade = false;
                                        bool isRepair = false;
                                        string installedDateString = "";
                                        if (!string.IsNullOrEmpty(CollectionName)) {
                                            if (!string.IsNullOrEmpty(CollectionGUID)) {
                                                var cs = cp.CSNew();
                                                var installedDate = DateTime.MinValue;
                                                bool IsOnSite = false;
                                                if (cs.Open("Add-on Collections", "ccguid=" + cp.Db.EncodeSQLText(CollectionGUID), "", true, "ID,ModifiedDate")) {
                                                    installedDate = cs.GetDate("ModifiedDate");
                                                    installedDateString = installedDate.ToShortDateString();
                                                    IsOnSite = true;
                                                }
                                                cs.Close();
                                                CollectionCheckbox = "<input TYPE=\"CheckBox\" NAME=\"LibraryRow\" VALUE=\"" + RowPtr + "\" onClick=\"clearLibraryRows('" + RowPtr + "');\">" + cp.Html.Hidden("LibraryRowGuid" + RowPtr, CollectionGUID) + cp.Html.Hidden("LibraryRowName" + RowPtr, CollectionName);
                                                if (!IsOnSite) {
                                                    // 
                                                    // -- not installed
                                                    isInstall = true;
                                                    showAddon = true;
                                                    CollectionCheckbox += "&nbsp;Install";
                                                } else if (installedDate >= CollectionModifiedDate) {
                                                    // 
                                                    // -- up to date, reinstall
                                                    isRepair = true;
                                                    showAddon = true;
                                                    CollectionCheckbox += "&nbsp;Reinstall";
                                                } else {
                                                    // 
                                                    // -- old version, upgrade
                                                    isUpgrade = true;
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
                                        bool requiresNewerContensive = false;
                                        if (!string.IsNullOrEmpty(CollectionContensiveVersion)) {
                                            try {
                                                var requiredVersion = new Version(CollectionContensiveVersion);
                                                var currentVersion = new Version(cp.Version);
                                                requiresNewerContensive = requiredVersion > currentVersion;
                                            } catch (Exception) {
                                                // -- ignore version parse errors
                                            }
                                        }
                                        if (showAddon) {
                                            viewModel.collectionList.Add(new LibraryViewModel_collectionList() {
                                                name = CollectionName,
                                                imageLink = CollectionImageLink,
                                                checkbox = CollectionCheckbox,
                                                lastUpdatedString = CollectionModifiedDateCaption,
                                                description = CollectionDescription,
                                                isInstall = isInstall,
                                                isUpgrade = isUpgrade,
                                                isRepair = isRepair,
                                                buttonValue = CollectionGUID,
                                                installedDate = installedDateString,
                                                requiresNewerContensive = requiresNewerContensive
                                            });
                                        }
                                        RowPtr += 1;
                                        break;
                                    }
                            }
                        }
                    }
                    string layout = cp.Layout.GetLayout(_Constants.guidAddonManagerLibraryListCell, _Constants.nameAddonManagerLibraryLisCell, _Constants.pathFilenameAddonManagerLibraryLisCell);
                    form.body = cp.Mustache.Render(layout, viewModel);
                    // 
                    // --------------------------------------------------------------------------------
                    // Build Page from tabs
                    // --------------------------------------------------------------------------------
                    // 
                    form.addFormButton(_Constants.ButtonCancel);
                    form.addFormButton(_Constants.ButtonOK);
                }
                // 
                // Output the Add-on
                // 
                form.title = "Add-on Library";
                form.description = "";
                if (!string.IsNullOrEmpty(status)) {
                    form.description = form.description + "<div style=\"Margin-left:50px\">" + status + "</div>";
                }
                return form.getHtml();
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
    /// <summary>
    /// rendering class for the Add-on Library
    /// </summary>
    public class LibraryViewModel {
        /// <summary>
        /// list of collections
        /// </summary>
        public List<LibraryViewModel_collectionList> collectionList { get; set; } = new List<LibraryViewModel_collectionList>();
    }
    /// <summary>
    /// a collection entry in the library list
    /// </summary>
    public class LibraryViewModel_collectionList {
        /// <summary>
        /// collection name
        /// </summary>
        public string name { get; set; }
        /// <summary>
        /// image
        /// </summary>
        public string imageLink { get; set; }
        /// <summary>
        /// checkbox (should be in the form)
        /// </summary>
        public string checkbox { get; set; }
        /// <summary>
        /// string that represents the date updated
        /// </summary>
        public string lastUpdatedString { get; set; }
        /// <summary>
        /// description
        /// </summary>
        public string description { get; set; }
        /// <summary>
        /// it is not currently installed
        /// </summary>
        public bool isInstall { get; set; }
        /// <summary>
        /// a previous version is installed
        /// </summary>
        public bool isUpgrade { get; set; }
        /// <summary>
        /// it is installed and up to date. Click to reinstall/Repair
        /// </summary>
        public bool isRepair { get; set; }
        /// <summary>
        /// The button name is 'button'
        /// This is the value of the button pressed. It will be the collectionId
        /// </summary>
        public string buttonValue { get; set; } = "";
        /// <summary>
        /// if isupgrade or isrepair, this is the date it was installed
        /// </summary>
        public string installedDate { get; set; } = "";
        /// <summary>
        /// true if the collection requires a newer version of Contensive than is currently installed
        /// </summary>
        public bool requiresNewerContensive { get; set; }

    }
}