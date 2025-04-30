using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.AddonManager51 {
    public sealed class InstallController {
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