# AJAX Collection Install Status Plan

## Overview

Replace the current form-POST-based collection install flow with AJAX-based install trigger and status polling. A single new remote method addon handles both install requests and status queries. On page load, detect already-running install tasks and resume polling.

## New Remote Method Addon

- **Name:** `addonManager-CollectionInstallRemote`
- **GUID:** `{8FD7D732-1AF7-42B9-884A-612B009BC100}`
- **DotNetClass:** `Contensive.AddonManager.Addons.CollectionInstallRemoteClass`
- **RemoteMethod:** Yes
- **AsAjax:** Yes

### Request/Response Contract

**Install request** (POST JSON or form-encoded):
- `action=install`
- `collectionGuid={guid}`

Response:
```json
{ "success": true, "taskId": 42, "collectionGuid": "{...}" }
```
or on error (e.g., version too old for background install):
```json
{ "success": false, "error": "message" }
```

**Status request:**
- `action=status`
- `taskId=42`

Response:
```json
{ "taskId": 42, "isQueued": false, "isRunning": true, "isCompleted": false, "isFailed": false, "failMessage": "" }
```

**Active tasks request** (page load — find running install tasks):
- `action=activeTasks`

Response:
```json
{ "tasks": [ { "taskId": 42, "collectionGuid": "{...}", "isRunning": true, "isQueued": false, "isCompleted": false, "isFailed": false } ] }
```

### Auth Guard

The remote method checks `cp.User.IsAdmin` and returns `{ "success": false, "error": "Administrator access required." }` if not admin.

## Implementation Steps

### Step 1: Add remote method addon to collection XML

Add the new `<Addon>` element to `collections/Add-on Manager/Add-on Manager.xml` with:
- `<RemoteMethod>Yes</RemoteMethod>`
- `<AsAjax>Yes</AsAjax>`
- `<DotNetClass>Contensive.AddonManager.Addons.CollectionInstallRemoteClass</DotNetClass>`

### Step 2: Add constants

Add to `Constants.cs`:
- `guidCollectionInstallRemote` — the remote method addon GUID
- `taskNamePrefix` — `"CollectionInstall:"` prefix used when naming tasks so we can find active ones

### Step 3: Create `CollectionInstallRemoteClass.cs`

New file: `source/addonManager51/Addons/CollectionInstallRemoteClass.cs`

Handles three actions:

**`install`:**
1. Validate `cp.User.IsAdmin`
2. Get `collectionGuid` from request
3. Call `cp.Addon.ExecuteAsProcessTask(Constants.guidCollectionInstallTask, args)` to get taskId
4. Store mapping: save `collectionGuid` in task argument so we can look it up later
5. Return JSON `{ success: true, taskId, collectionGuid }`
6. Fallback: if platform doesn't support `ExecuteAsProcessTask`, do synchronous install and return `{ success: true, taskId: 0, isCompleted: true }`

**`status`:**
1. Validate `cp.User.IsAdmin`
2. Get `taskId` from request
3. Call `cp.Task.GetStatus(taskId)` to get `TaskStatusModel`
4. Return JSON with status fields

**`activeTasks`:**
1. Validate `cp.User.IsAdmin`
2. Query `cctasks` for tasks where the `cmdDetail` contains `collectionGuid` argument AND `dateCompleted is null` (queued or running)
3. For each, call `cp.Task.GetStatus(taskId)` and include `collectionGuid` extracted from the task's `cmdDetail`
4. Return JSON array of active tasks

### Step 4: Update `CollectionLibraryClass.cs`

**Remove form-POST install handling:**
- Remove the button-processing block that calls `installOrQueueCollection()` for Install/Upgrade/Reinstall/UpgradeAll
- Keep the `buttonCancel` handling (still a form POST redirect)
- Keep the `buttonRemove` (uninstall) handling (still form POST, it's synchronous)
- The form body still renders the collection cards, but buttons change from `type="submit"` to `type="button"` with `onclick` handlers

**Add `collectionGuid` to the view model items:**
- `LibraryViewModel_collectionList` already has `buttonValue` which holds the collection GUID — this is sufficient

### Step 5: Update Mustache template (`AddonManagerLibraryBody.html`)

**Change Install/Upgrade/Reinstall buttons:**
- Change from `type="submit" name="button" value="{{buttonValue}}"` to `type="button" onclick="amInstallCollection('{{buttonValue}}', this)"`
- Add a `data-collection-guid="{{buttonValue}}"` attribute to each card for lookup
- Add a status element inside each card: `<span class="am-install-status" data-collection-guid="{{buttonValue}}"></span>`

**Keep these as form submits:**
- Cancel button (form-level, stays as-is)
- Remove button (stays as `type="submit"`)

**Add JavaScript at the bottom:**

```javascript
// -- remote method URL
var amRemoteUrl = '/addonManager-CollectionInstallRemote';

// -- install a collection via AJAX
function amInstallCollection(collectionGuid, btn) {
    btn.disabled = true;
    var statusEl = document.querySelector('.am-install-status[data-collection-guid="' + collectionGuid + '"]');
    if (statusEl) statusEl.innerHTML = 'Queuing...';

    fetch(amRemoteUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: 'action=install&collectionGuid=' + encodeURIComponent(collectionGuid)
    })
    .then(r => r.json())
    .then(data => {
        if (data.success && data.taskId > 0) {
            if (statusEl) statusEl.innerHTML = 'Installing...';
            amPollStatus(data.taskId, collectionGuid, btn);
        } else if (data.success && data.taskId === 0) {
            // synchronous install completed
            if (statusEl) statusEl.innerHTML = 'Installed';
            setTimeout(() => location.reload(), 1500);
        } else {
            if (statusEl) statusEl.innerHTML = 'Error: ' + (data.error || 'Unknown');
            btn.disabled = false;
        }
    })
    .catch(err => {
        if (statusEl) statusEl.innerHTML = 'Request failed';
        btn.disabled = false;
    });
}

// -- poll task status
function amPollStatus(taskId, collectionGuid, btn) {
    fetch(amRemoteUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: 'action=status&taskId=' + taskId
    })
    .then(r => r.json())
    .then(data => {
        var statusEl = document.querySelector('.am-install-status[data-collection-guid="' + collectionGuid + '"]');
        if (data.isCompleted) {
            if (statusEl) statusEl.innerHTML = 'Installed';
            setTimeout(() => location.reload(), 1500);
        } else if (data.isFailed) {
            if (statusEl) statusEl.innerHTML = 'Failed: ' + (data.failMessage || 'Unknown error');
            if (btn) btn.disabled = false;
        } else {
            // still running or queued
            var msg = data.isRunning ? 'Installing...' : 'Queued...';
            if (statusEl) statusEl.innerHTML = msg + ' <img src="https://s3.amazonaws.com/cdn.contensive.com/assets/20201227/images/ajax-loader-small.gif" width="16" height="16">';
            setTimeout(() => amPollStatus(taskId, collectionGuid, btn), 3000);
        }
    })
    .catch(() => {
        setTimeout(() => amPollStatus(taskId, collectionGuid, btn), 5000);
    });
}

// -- on page load, check for active install tasks
function amCheckActiveTasks() {
    fetch(amRemoteUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: 'action=activeTasks'
    })
    .then(r => r.json())
    .then(data => {
        if (data.tasks && data.tasks.length > 0) {
            data.tasks.forEach(t => {
                var btn = document.querySelector('[data-collection-guid="' + t.collectionGuid + '"] button[type="button"]');
                if (btn) btn.disabled = true;
                var statusEl = document.querySelector('.am-install-status[data-collection-guid="' + t.collectionGuid + '"]');
                if (statusEl) {
                    statusEl.innerHTML = (t.isRunning ? 'Installing...' : 'Queued...')
                        + ' <img src="https://s3.amazonaws.com/cdn.contensive.com/assets/20201227/images/ajax-loader-small.gif" width="16" height="16">';
                }
                amPollStatus(t.taskId, t.collectionGuid, btn);
            });
        }
    })
    .catch(() => { /* silent */ });
}

amCheckActiveTasks();
```

### Step 6: Handle Upgrade All via AJAX

The "Upgrade All" button changes from a form submit to a JavaScript function:
- Loop through all upgrade GUIDs from the hidden `upgradeAllGuids` field
- Call `amInstallCollection()` for each one sequentially (or in parallel — each gets its own card status)
- Remove the `buttonUpgradeAll` form button; add a regular button with `onclick="amUpgradeAll()"`

### Step 7: Version fallback

In `CollectionInstallRemoteClass.cs`, if `ExecuteAsProcessTask` is not available (older Contensive), catch `MissingMethodException` and fall back to synchronous `InstallCollectionFromLibrary()`, returning `{ success: true, taskId: 0, isCompleted: true }`.

## Files Modified

| File | Change |
|------|--------|
| `collections/Add-on Manager/Add-on Manager.xml` | Add remote method addon element |
| `source/addonManager51/Constants.cs` | Add remote addon GUID constant |
| `source/addonManager51/Addons/CollectionInstallRemoteClass.cs` | **NEW** — remote method handler |
| `source/addonManager51/Addons/CollectionLibraryClass.cs` | Remove form-POST install logic, keep Remove/Cancel |
| `ui/layoutFiles/AddonManagerLibraryBody.html` | AJAX buttons, status elements, JS polling |

## Notes

- The `Upgrade All` button on the layout builder form buttons bar will be removed. Instead, an inline "Upgrade All" button will appear in the filter bar area of the template, wired to JS.
- The `OK` button is legacy and can be removed since individual Install/Upgrade/Reinstall buttons handle actions.
- The `Cancel` button stays as a form POST.
- The `Remove` buttons stay as form POSTs (uninstall is synchronous).
- For `activeTasks`, we query `cctasks` directly via `cp.Db` for tasks where the addon GUID matches `guidCollectionInstallTask` and `dateCompleted is null`. The `cmdDetail` XML contains the argument list including `collectionGuid`.
