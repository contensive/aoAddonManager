using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.AddonManager51 {
    /// <summary>
    /// 
    /// </summary>
    public sealed class InstallController {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="TargetCollectionID"></param>
        public static  void UninstallCollection(CPBaseClass cp, int TargetCollectionID) {
            string TargetCollectionName = cp.Content.GetRecordName("Add-on Collections", TargetCollectionID);
            // 
            // Clean up rules associating this collection to other objects
            // 
            cp.Content.Delete("Add-on Collection CDef Rules", "collectionid=" + TargetCollectionID);
            cp.Content.Delete("Add-on Collection Module Rules", "collectionid=" + TargetCollectionID);
            var cs = cp.CSNew();
            // 
            // Delete any addons from this collection
            // 
            if (cs.Open("add-ons", "collectionid=" + TargetCollectionID)) {
                do {
                    // 
                    // Clean up the rules that might have pointed to the addon
                    // 
                    int addonid = cs.GetInteger("id");
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
                int AddonNavigatorID = 0;
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

        // 
        // -- method provided here because these methods are not included in the c41 interface, so this call can only be created if v5 code
        public static bool installCollectionFromLibrary(CPBaseClass cp, string collectionGuid, ref string ErrorMessage) {
            // 
            cp.Utils.AppendLog("installCollectionFromLibrary, collectionGuid [" + collectionGuid + "]");
            // 
            cp.Addon.InstallCollectionFromLibrary(collectionGuid, ref ErrorMessage);
            return true;
        }
        // 
        // -- method provided here because these methods are not included in the c41 interface, so this call can only be created if v5 code
        public static bool installCollectionFromFolder(CPBaseClass cp, string privatePathFilename, ref string ErrorMessage) {
            try {
                // 
                cp.Utils.AppendLog("installCollectionFromFolder, privatePathFilename [" + privatePathFilename + "]");
                // 
                return cp.Addon.InstallCollectionFile(privatePathFilename, ref ErrorMessage);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // -- method provided here because these methods are not included in the c41 interface, so this call can only be created if v5 code
        public static bool installCollectionFromUpload(CPBaseClass cp, string requestName, ref string ErrorMessage) {
            // 
            cp.Utils.AppendLog("installCollectionFromUpload, requestName [" + requestName + "]");
            try {
                // 
                string privatePath = "CollectionUpload" + cp.Utils.CreateGuid().Replace("{", "").Replace("-", "").Replace("}", "") + @"\";
                string uploadFilename = "";
                bool result = false;
                if (cp.PrivateFiles.SaveUpload(requestName, privatePath, ref uploadFilename)) {
                    result = installCollectionFromFolder(cp, privatePath + uploadFilename, ref ErrorMessage);
                }
                cp.PrivateFiles.DeleteFolder(privatePath);
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}