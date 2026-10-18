using Contensive.BaseClasses;
using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace Contensive.AddonManager.Addons {
    /// <summary>
    /// Remote method addon that handles AJAX collection install trigger and status polling.
    /// Single endpoint handles three actions: install, status, and activeTasks.
    /// Uses reflection for ExecuteAsProcessTask and cp.Task.GetStatus so the addon compiles
    /// against older CPBase packages and gracefully falls back when methods are unavailable.
    /// </summary>
    public class CollectionInstallRemoteClass : AddonBaseClass {
        /// <summary>
        /// Addon entry point — routes AJAX requests to install, status, or activeTasks handlers
        /// </summary>
        public override object Execute(CPBaseClass cp) {
            try {
                if (!cp.User.IsAdmin) {
                    return jsonError("Administrator access required.");
                }
                string action = cp.Doc.GetText("action");
                switch (action.ToLowerInvariant()) {
                    case "install":
                        return handleInstall(cp);
                    case "status":
                        return handleStatus(cp);
                    case "activetasks":
                        return handleActiveTasks(cp);
                    default:
                        return jsonError("Unknown action.");
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return jsonError($"Unexpected error: {ex.Message}");
            }
        }
        //
        /// <summary>
        /// Handle install action: queue a collection install as a background task and return the task id.
        /// Uses reflection to call ExecuteAsProcessTask. Falls back to synchronous install.
        /// </summary>
        private static string handleInstall(CPBaseClass cp) {
            string collectionGuid = cp.Doc.GetText("collectionGuid");
            if (string.IsNullOrEmpty(collectionGuid)) {
                return jsonError("collectionGuid is required.");
            }
            //
            // -- try background install with task id tracking via reflection
            //
            int taskId = invokeExecuteAsProcessTask(cp, Constants.guidCollectionInstallTask, new Dictionary<string, string> { { "collectionGuid", collectionGuid } });
            if (taskId > 0) {
                return $"{{\"success\":true,\"taskId\":{taskId},\"collectionGuid\":\"{jsonEscape(collectionGuid)}\"}}";
            }
            //
            // -- fallback: try fire-and-forget ExecuteAsProcess (returns void, no task id)
            //
            try {
                var args = new Dictionary<string, string> { { "collectionGuid", collectionGuid } };
                cp.Addon.ExecuteAsProcess(Constants.guidCollectionInstallTask, args);
                //
                // -- queued but we don't have a task id, return taskId=-1 to signal "no polling"
                //
                return $"{{\"success\":true,\"taskId\":-1,\"collectionGuid\":\"{jsonEscape(collectionGuid)}\"}}";
            } catch (Exception) {
                //
                // -- last resort: synchronous install
                //
                try {
                    string errorMessage = "";
                    cp.Addon.InstallCollectionFromLibrary(collectionGuid, ref errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage)) {
                        return $"{{\"success\":true,\"taskId\":0,\"isCompleted\":true,\"warning\":\"{jsonEscape(errorMessage)}\"}}";
                    }
                    return "{\"success\":true,\"taskId\":0,\"isCompleted\":true}";
                } catch (Exception exSync) {
                    return jsonError($"Install failed: {exSync.Message}");
                }
            }
        }
        //
        /// <summary>
        /// Handle status action: return the current status of a task by id.
        /// Tries cp.Task.GetStatus via reflection, falls back to direct DB query.
        /// </summary>
        private static string handleStatus(CPBaseClass cp) {
            int taskId = cp.Doc.GetInteger("taskId");
            if (taskId <= 0) {
                return jsonError("taskId is required.");
            }
            //
            // -- try unified task API via reflection
            //
            string result = invokeGetStatus(cp, taskId);
            if (result != null) {
                return result;
            }
            //
            // -- fallback to direct DB query
            //
            return getStatusFromDb(cp, taskId);
        }
        //
        /// <summary>
        /// Handle activeTasks action: find all currently running or queued collection install tasks.
        /// Queries the cctasks table directly for tasks matching the CollectionInstall addon GUID.
        /// </summary>
        private static string handleActiveTasks(CPBaseClass cp) {
            try {
                string addonGuid = Constants.guidCollectionInstallTask;
                string sql = $"select id, cmdDetail from cctasks where (dateCompleted is null) and (cmdDetail like '%{addonGuid}%')";
                using (DataTable dt = cp.Db.ExecuteQuery(sql)) {
                    if (dt == null || dt.Rows.Count == 0) {
                        return "{\"tasks\":[]}";
                    }
                    var taskEntries = new List<string>();
                    foreach (DataRow row in dt.Rows) {
                        int taskId = 0;
                        if (row["id"] != null && row["id"] != DBNull.Value) {
                            int.TryParse(row["id"].ToString(), out taskId);
                        }
                        string cmdDetail = row["cmdDetail"]?.ToString() ?? "";
                        string collectionGuid = extractCollectionGuid(cmdDetail);
                        string statusJson = getTaskStatusJsonFromDb(cp, taskId, collectionGuid);
                        taskEntries.Add(statusJson);
                    }
                    return $"{{\"tasks\":[{string.Join(",", taskEntries)}]}}";
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return "{\"tasks\":[]}";
            }
        }
        //
        // ==============================================================================================
        // Reflection helpers — call methods that may not exist in the compiled-against CPBase version
        // ==============================================================================================
        //
        /// <summary>
        /// Call cp.Addon.ExecuteAsProcessTask(guid, args) via reflection. Returns task id or 0 on failure.
        /// </summary>
        private static int invokeExecuteAsProcessTask(CPBaseClass cp, string addonGuid, Dictionary<string, string> args) {
            try {
                MethodInfo method = cp.Addon.GetType().GetMethod(
                    "ExecuteAsProcessTask",
                    new Type[] { typeof(string), typeof(Dictionary<string, string>) }
                );
                if (method == null) { return 0; }
                object result = method.Invoke(cp.Addon, new object[] { addonGuid, args });
                if (result is int intResult) { return intResult; }
                return 0;
            } catch (Exception) {
                return 0;
            }
        }
        //
        /// <summary>
        /// Call cp.Task.GetStatus(taskId) via reflection. Returns JSON string or null if unavailable.
        /// </summary>
        private static string invokeGetStatus(CPBaseClass cp, int taskId) {
            try {
                //
                // -- get the Task property
                //
                PropertyInfo taskProp = cp.GetType().GetProperty("Task");
                if (taskProp == null) { return null; }
                object taskObj = taskProp.GetValue(cp);
                if (taskObj == null) { return null; }
                //
                // -- call GetStatus(int taskId)
                //
                MethodInfo getStatusMethod = taskObj.GetType().GetMethod("GetStatus", new Type[] { typeof(int) });
                if (getStatusMethod == null) { return null; }
                object statusObj = getStatusMethod.Invoke(taskObj, new object[] { taskId });
                if (statusObj == null) {
                    //
                    // -- task not found (deleted by housekeeping), treat as completed
                    //
                    return $"{{\"taskId\":{taskId},\"isQueued\":false,\"isRunning\":false,\"isCompleted\":true,\"isFailed\":false,\"failMessage\":\"\"}}";
                }
                //
                // -- read properties from TaskStatusModel via reflection
                //
                Type statusType = statusObj.GetType();
                bool isQueued = (bool)(statusType.GetProperty("isQueued")?.GetValue(statusObj) ?? false);
                bool isRunning = (bool)(statusType.GetProperty("isRunning")?.GetValue(statusObj) ?? false);
                bool isCompleted = (bool)(statusType.GetProperty("isCompleted")?.GetValue(statusObj) ?? false);
                bool isFailed = (bool)(statusType.GetProperty("isFailed")?.GetValue(statusObj) ?? false);
                string failMessage = statusType.GetProperty("failMessage")?.GetValue(statusObj)?.ToString() ?? "";
                return $"{{\"taskId\":{taskId},\"isQueued\":{boolJson(isQueued)},\"isRunning\":{boolJson(isRunning)},\"isCompleted\":{boolJson(isCompleted)},\"isFailed\":{boolJson(isFailed)},\"failMessage\":\"{jsonEscape(failMessage)}\"}}";
            } catch (Exception) {
                return null;
            }
        }
        //
        // ==============================================================================================
        // DB fallback helpers
        // ==============================================================================================
        //
        /// <summary>
        /// Fallback: get task status directly from DB when cp.Task is not available.
        /// </summary>
        private static string getStatusFromDb(CPBaseClass cp, int taskId) {
            try {
                string sql = $"select dateStarted, dateCompleted, failMessage from cctasks where id={taskId}";
                using (DataTable dt = cp.Db.ExecuteQuery(sql)) {
                    if (dt == null || dt.Rows.Count == 0) {
                        return $"{{\"taskId\":{taskId},\"isQueued\":false,\"isRunning\":false,\"isCompleted\":true,\"isFailed\":false,\"failMessage\":\"\"}}";
                    }
                    DataRow row = dt.Rows[0];
                    bool hasStarted = row["dateStarted"] != null && row["dateStarted"] != DBNull.Value;
                    bool hasCompleted = row["dateCompleted"] != null && row["dateCompleted"] != DBNull.Value;
                    string failMessage = row["failMessage"]?.ToString() ?? "";
                    bool isQueued = !hasStarted;
                    bool isRunning = hasStarted && !hasCompleted;
                    bool isCompleted = hasCompleted && string.IsNullOrEmpty(failMessage);
                    bool isFailed = hasCompleted && !string.IsNullOrEmpty(failMessage);
                    return $"{{\"taskId\":{taskId},\"isQueued\":{boolJson(isQueued)},\"isRunning\":{boolJson(isRunning)},\"isCompleted\":{boolJson(isCompleted)},\"isFailed\":{boolJson(isFailed)},\"failMessage\":\"{jsonEscape(failMessage)}\"}}";
                }
            } catch (Exception) {
                return $"{{\"taskId\":{taskId},\"isQueued\":false,\"isRunning\":false,\"isCompleted\":true,\"isFailed\":false,\"failMessage\":\"\"}}";
            }
        }
        //
        /// <summary>
        /// Get task status from DB with collectionGuid included in response.
        /// </summary>
        private static string getTaskStatusJsonFromDb(CPBaseClass cp, int taskId, string collectionGuid) {
            try {
                string sql = $"select dateStarted, dateCompleted from cctasks where id={taskId}";
                using (DataTable dt = cp.Db.ExecuteQuery(sql)) {
                    if (dt == null || dt.Rows.Count == 0) {
                        return $"{{\"taskId\":{taskId},\"collectionGuid\":\"{jsonEscape(collectionGuid)}\",\"isQueued\":false,\"isRunning\":false,\"isCompleted\":true,\"isFailed\":false}}";
                    }
                    DataRow row = dt.Rows[0];
                    bool hasStarted = row["dateStarted"] != null && row["dateStarted"] != DBNull.Value;
                    bool hasCompleted = row["dateCompleted"] != null && row["dateCompleted"] != DBNull.Value;
                    bool isQueued = !hasStarted;
                    bool isRunning = hasStarted && !hasCompleted;
                    bool isCompleted = hasCompleted;
                    return $"{{\"taskId\":{taskId},\"collectionGuid\":\"{jsonEscape(collectionGuid)}\",\"isQueued\":{boolJson(isQueued)},\"isRunning\":{boolJson(isRunning)},\"isCompleted\":{boolJson(isCompleted)},\"isFailed\":false}}";
                }
            } catch (Exception) {
                return $"{{\"taskId\":{taskId},\"collectionGuid\":\"{jsonEscape(collectionGuid)}\",\"isQueued\":false,\"isRunning\":false,\"isCompleted\":true,\"isFailed\":false}}";
            }
        }
        //
        // ==============================================================================================
        // cmdDetail XML parsing
        // ==============================================================================================
        //
        /// <summary>
        /// Extract collectionGuid from the task's cmdDetail XML.
        /// The cmdDetail contains argument key-value pairs; we look for collectionGuid.
        /// </summary>
        private static string extractCollectionGuid(string cmdDetail) {
            if (string.IsNullOrEmpty(cmdDetail)) { return ""; }
            //
            // -- cmdDetail XML like: <cmd><row key="collectionGuid" value="{guid}"/></cmd>
            //
            string searchKey = "collectionGuid";
            int keyIdx = cmdDetail.IndexOf(searchKey, StringComparison.OrdinalIgnoreCase);
            if (keyIdx < 0) { return ""; }
            //
            // -- find value=" after the key
            //
            int valueIdx = cmdDetail.IndexOf("value=\"", keyIdx, StringComparison.OrdinalIgnoreCase);
            if (valueIdx < 0) {
                valueIdx = cmdDetail.IndexOf("value=&quot;", keyIdx, StringComparison.OrdinalIgnoreCase);
                if (valueIdx < 0) { return ""; }
                valueIdx += "value=&quot;".Length;
                int endIdx = cmdDetail.IndexOf("&quot;", valueIdx, StringComparison.OrdinalIgnoreCase);
                return endIdx < 0 ? "" : cmdDetail.Substring(valueIdx, endIdx - valueIdx);
            }
            valueIdx += "value=\"".Length;
            int closeIdx = cmdDetail.IndexOf("\"", valueIdx, StringComparison.Ordinal);
            return closeIdx < 0 ? "" : cmdDetail.Substring(valueIdx, closeIdx - valueIdx);
        }
        //
        // ==============================================================================================
        // JSON helpers
        // ==============================================================================================
        //
        private static string jsonError(string message) {
            return $"{{\"success\":false,\"error\":\"{jsonEscape(message)}\"}}";
        }
        //
        private static string jsonEscape(string s) {
            if (string.IsNullOrEmpty(s)) { return ""; }
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n");
        }
        //
        private static string boolJson(bool value) {
            return value ? "true" : "false";
        }
    }
}
