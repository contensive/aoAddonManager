using Contensive.AddonManager.Controllers;
using Contensive.BaseClasses;
using Contensive.Models.Db;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace Contensive.AddonManager.Addons {
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
                return "<p>An error occurred loading the collection library. Please try again.</p>";
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
                string Button = cp.Doc.GetText(Constants.requestNameButton);
                string ButtonRemove = cp.Doc.GetText(Constants.buttonRemove);
                if ((Button ?? "") == Constants.buttonCancel) {
                    //
                    // ----- redirect back to the root
                    //
                    cp.Response.Redirect(cp.Site.GetText("adminUrl"));
                    return "";
                }
                string status = "";
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    string CollectionGUID = "";
                    if (!string.IsNullOrEmpty(ButtonRemove ?? "")) {
                        //
                        // -- uninstall (synchronous, stays as form POST)
                        //
                        var collection = DbBaseModel.create<AddonCollectionModel>(cp, ButtonRemove);
                        InstallController.uninstallCollection(cp, collection.id);
                    }


                    //
                    System.Xml.XmlDocument LibCollections = new();
                    LibCollections.Load($"https://support.contensive.com/GetCollectionList?iv={cp.Version}&key={cp.Utils.EncodeRequestVariable(SiteKey)}&name={cp.Utils.EncodeRequestVariable(cp.Site.Name)}&primaryDomain={cp.Utils.EncodeRequestVariable(cp.Site.DomainPrimary)}");
                    //
                    LibraryViewModel viewModel = new() {
                        filterSearch = cp.Doc.GetText("amFilterSearchValue"),
                        filterInstalled = cp.Doc.GetBoolean("amFilterInstalledValue"),
                        filterUpgrades = cp.Doc.GetBoolean("amFilterUpgradesValue"),
                        filterNotInstalled = cp.Doc.GetBoolean("amFilterNotInstalledValue"),
                        remoteMethodUrl = "/addonManager-CollectionInstallRemote"
                    };
                    if ((LibCollections.DocumentElement.Name.ToLower() ?? "") != (Constants.collectionListRootNode.ToLower() ?? "")) {
                        string UserError = $"There was an error reading the Collection Library file. The '{Constants.collectionListRootNode}' element was not found.";
                        status += $"<BR>{UserError}";
                        cp.UserError.Add(UserError);
                    } else {
                        var RowPtr = default(int);
                        // 
                        // Go through file to validate the XML, and build status message -- since service process can not communicate to user
                        // 
                        RowPtr = 0;
                        // Content = ""
                        foreach (System.Xml.XmlNode CDef_Node in LibCollections.DocumentElement.ChildNodes) {
                            string CollectionImageLink = "";
                            string CollectionName = "";
                            var CollectionModifiedDate = DateTime.MinValue;
                            string CollectionModifiedDateCaption = "";
                            string CollectionDescription = "";
                            string CollectionHelpLink = "";
                            string CollectionDemoLink = "";
                            string CollectionContensiveVersion = "";
                            switch (CDef_Node.Name.ToLower() ?? "") {
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
                                                case "lastmodifieddate": {
                                                        // 
                                                        // last modified
                                                        // 
                                                        CollectionModifiedDate = DateTime.MinValue;
                                                        if (DateTime.TryParse(CollectionNode.InnerText, out DateTime parsedDate)) {
                                                            CollectionModifiedDate = parsedDate;
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
                                        bool isNotUpdatable = false;
                                        string installedDateString = "";
                                        if (!string.IsNullOrEmpty(CollectionName)) {
                                            if (!string.IsNullOrEmpty(CollectionGUID)) {
                                                var cs = cp.CSNew();
                                                var installedDate = DateTime.MinValue;
                                                bool IsOnSite = false;
                                                if (cs.Open("Add-on Collections", $"ccguid={cp.Db.EncodeSQLText(CollectionGUID)}", "", true, "ID,ModifiedDate,updatable")) {
                                                    installedDate = cs.GetDate("ModifiedDate");
                                                    installedDateString = installedDate.ToShortDateString();
                                                    IsOnSite = true;
                                                    isNotUpdatable = !cs.GetBoolean("updatable");
                                                }
                                                cs.Close();
                                                if (!IsOnSite) {
                                                    //
                                                    // -- not installed
                                                    isInstall = true;
                                                    showAddon = true;
                                                } else if (installedDate >= CollectionModifiedDate) {
                                                    //
                                                    // -- up to date, reinstall
                                                    isRepair = true;
                                                    showAddon = true;
                                                } else {
                                                    //
                                                    // -- old version, upgrade
                                                    isUpgrade = true;
                                                    showAddon = true;
                                                }
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(CollectionDemoLink)) {
                                            CollectionDescription += $"<div class=\"amDemoLink\"><a target=\"_blank\" href=\"{CollectionDemoLink}\">Demo</a></div>";
                                        }
                                        if (!string.IsNullOrEmpty(CollectionHelpLink)) {
                                            CollectionDescription += $"<div class=\"amHelpLink\"><a target=\"_blank\" href=\"{CollectionHelpLink}\">Reference</a></div>";
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
                                                lastUpdatedString = CollectionModifiedDateCaption,
                                                description = CollectionDescription,
                                                isInstall = isInstall,
                                                isUpgrade = isUpgrade,
                                                isRepair = isRepair,
                                                buttonValue = CollectionGUID,
                                                installedDate = installedDateString,
                                                requiresNewerContensive = requiresNewerContensive,
                                                notUpdatable = isNotUpdatable
                                            });
                                        }
                                        RowPtr += 1;
                                        break;
                                    }
                            }
                        }
                    }
                    //
                    // -- build a comma-separated list of GUIDs for all upgradable collections
                    //
                    string upgradeAllGuids = string.Join(",", viewModel.collectionList
                        .Where(c => c.isUpgrade && !c.notUpdatable)
                        .Select(c => c.buttonValue));
                    viewModel.upgradeAllGuids = upgradeAllGuids;
                    viewModel.hasUpgrades = !string.IsNullOrEmpty(upgradeAllGuids);
                    //
                    string layout = cp.Layout.GetLayout(Constants.guidAddonManagerLibraryListCell, Constants.nameAddonManagerLibraryLisCell, Constants.pathFilenameAddonManagerLibraryLisCell);
                    form.body = cp.Mustache.Render(layout, viewModel);
                    //
                    // --------------------------------------------------------------------------------
                    // Build Page from tabs
                    // --------------------------------------------------------------------------------
                    //
                    form.addFormButton(Constants.buttonCancel);
                }
                // 
                // Output the Add-on
                // 
                form.title = "Add-on Library";
                form.description = "";
                if (!string.IsNullOrEmpty(status)) {
                    form.description += $"<div style=\"Margin-left:50px\">{status}</div>";
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
        public List<LibraryViewModel_collectionList> collectionList { get; set; } = [];
        /// <summary>
        /// persisted search text filter
        /// </summary>
        public string filterSearch { get; set; } = "";
        /// <summary>
        /// persisted installed filter checkbox
        /// </summary>
        public bool filterInstalled { get; set; }
        /// <summary>
        /// persisted upgrades filter checkbox
        /// </summary>
        public bool filterUpgrades { get; set; }
        /// <summary>
        /// persisted not-installed filter checkbox
        /// </summary>
        public bool filterNotInstalled { get; set; }
        /// <summary>
        /// comma-separated GUIDs of all collections that need upgrading (for Upgrade All button)
        /// </summary>
        public string upgradeAllGuids { get; set; } = "";
        /// <summary>
        /// true if there are collections that need upgrading
        /// </summary>
        public bool hasUpgrades { get; set; }
        /// <summary>
        /// URL of the remote method addon for AJAX install and status polling
        /// </summary>
        public string remoteMethodUrl { get; set; } = "";
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
        /// <summary>
        /// true if the installed collection is marked as not updatable
        /// </summary>
        public bool notUpdatable { get; set; }

    }
}