using System;
using Contensive.BaseClasses;
using Microsoft.VisualBasic;

namespace Contensive.Addons.AddonManager51 {
    static class GenericController {
        // 
        // 
        // 
        public static void getForm_AddonManager_DeleteNavigatorBranch(CPBaseClass cp, string EntryName, int EntryParentID) {
            try {
                // 
                var cs = cp.CSNew();
                var EntryID = default(int);
                // 
                if (EntryParentID == 0) {
                    cs.Open("Navigator Entries", "(name=" + cp.Db.EncodeSQLText(EntryName) + ")and((parentID is null)or(parentid=0))");
                } else {
                    cs.Open("Navigator Entries", "(name=" + cp.Db.EncodeSQLText(EntryName) + ")and(parentID=" + cp.Db.EncodeSQLNumber(EntryParentID) + ")");
                }
                if (cs.OK()) {
                    EntryID = cs.GetInteger("ID");
                }
                cs.Close();
                // 
                if (EntryID != 0) {
                    cs.Open("Navigator Entries", "(parentID=" + cp.Db.EncodeSQLNumber(EntryID) + ")");
                    while (cs.OK()) {
                        getForm_AddonManager_DeleteNavigatorBranch(cp, cs.GetText("name"), EntryID);
                        cs.GoNext();
                    }
                    cs.Close();
                    cp.Content.Delete("Navigator Entries", "id=" + EntryID);
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // 
        // 
        public static int getParentIDFromNameSpace(CPBaseClass cp, string ContentName, string NameSpacex) {
            int result = 0;
            try {
                string ParentNameSpace;
                string ParentName;
                int ParentID;
                int Pos;
                var cs = cp.CSNew();
                // 
                if (!string.IsNullOrEmpty(NameSpacex)) {
                    Pos = Strings.InStr(1, NameSpacex, ".");
                    if (Pos == 0) {
                        ParentName = NameSpacex;
                        ParentNameSpace = "";
                    } else {
                        ParentName = Strings.Mid(NameSpacex, Pos + 1);
                        ParentNameSpace = Strings.Mid(NameSpacex, 1, Pos - 1);
                    }
                    if (string.IsNullOrEmpty(ParentNameSpace)) {
                        if (cs.Open(ContentName, "(name=" + cp.Db.EncodeSQLText(ParentName) + ")and((parentid is null)or(parentid=0))", "ID", true, "ID")) {
                            result = cs.GetInteger("ID");
                        }
                        cs.Close();
                    } else {
                        ParentID = getParentIDFromNameSpace(cp, ContentName, ParentNameSpace);
                        if (cs.Open(ContentName, "(name=" + cp.Db.EncodeSQLText(ParentName) + ")and(parentid=" + ParentID + ")", "ID", true, "ID")) {
                            result = cs.GetInteger("ID");
                        }
                        cs.Close();
                    }
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return result;
        }
    }
}