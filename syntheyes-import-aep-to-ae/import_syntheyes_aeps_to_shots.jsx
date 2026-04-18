/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

// ─────────────────────────────────────────────────────────────────────────────
// import_syntheyes_aeps_to_shots.jsx
//
// Imports SynthEyes-exported .aep files into the current project and wires
// each one into its matching shot folder.
//
// For each imported .aep the script:
//   1. Matches the shot name (e.g. "VS_002") against folders inside "Shots"
//   2. Moves the imported folder into that shot folder
//   3. Copies all layers from "Camera01_3D" into "{shot}_comp"
//      pasted at the start of the shot clip already inside the target comp
//      (the layer whose name starts with the shot name, e.g. "VS_002…")
//
// HOW TO USE:
//   File > Scripts > Run Script File… → pick this file
// ─────────────────────────────────────────────────────────────────────────────

(function importSyntheyesAEPs() {

    // ── Settings ──────────────────────────────────────────────────────────────
    var SHOTS_FOLDER_NAME  = "Shots";        // top-level folder in the project
    var CAMERA_COMP_NAME   = "Camera01_3D";  // comp to copy layers from
    var TARGET_COMP_SUFFIX = "_comp";        // e.g. VS_002 + "_comp" = VS_002_comp

    // ── Helpers ───────────────────────────────────────────────────────────────

    // Find a direct child of a FolderItem by name (optionally filter by type)
    function findInFolder(folder, name, type) {
        for (var i = 1; i <= folder.numItems; i++) {
            var item = folder.item(i);
            if (item.name === name && (!type || item instanceof type)) {
                return item;
            }
        }
        return null;
    }

    // Recursively search a FolderItem for a CompItem by name
    function findCompInFolder(folder, name) {
        for (var i = 1; i <= folder.numItems; i++) {
            var item = folder.item(i);
            if (item instanceof CompItem && item.name === name) return item;
            if (item instanceof FolderItem) {
                var found = findCompInFolder(item, name);
                if (found) return found;
            }
        }
        return null;
    }

    // Collect all FolderItems that are direct children of a folder
    function childFolders(folder) {
        var result = [];
        for (var i = 1; i <= folder.numItems; i++) {
            if (folder.item(i) instanceof FolderItem) result.push(folder.item(i));
        }
        return result;
    }

    // Given a filename and a list of known shot names, return the first match
    // where the filename starts with that shot name (e.g. "VS_002_comp..." → "VS_002")
    function matchShotName(filename, shotNames) {
        // Sort longest-first so "VS_0021" doesn't match before "VS_002"
        var sorted = shotNames.slice().sort(function(a, b) { return b.length - a.length; });
        for (var i = 0; i < sorted.length; i++) {
            if (filename.indexOf(sorted[i]) === 0) return sorted[i];
        }
        return null;
    }

    // ── Sanity-check: we need an open project ─────────────────────────────────
    if (!app.project) {
        alert("No project is open. Please open your main project first.");
        return;
    }

    // ── Find the Shots folder ─────────────────────────────────────────────────
    var shotsFolder = null;
    for (var i = 1; i <= app.project.numItems; i++) {
        var it = app.project.item(i);
        if (it instanceof FolderItem && it.name === SHOTS_FOLDER_NAME) {
            shotsFolder = it;
            break;
        }
    }
    if (!shotsFolder) {
        alert("Could not find a folder named \"" + SHOTS_FOLDER_NAME + "\" in the project.\n"
            + "Check the SHOTS_FOLDER_NAME setting at the top of this script.");
        return;
    }

    // Collect known shot names from the Shots folder
    var shotFolders = childFolders(shotsFolder);
    if (shotFolders.length === 0) {
        alert("The \"" + SHOTS_FOLDER_NAME + "\" folder has no sub-folders.");
        return;
    }
    var shotNames = [];
    for (var s = 0; s < shotFolders.length; s++) shotNames.push(shotFolders[s].name);

    // ── Pick .aep files ───────────────────────────────────────────────────────
    // Use a function filter — macOS ignores the "Type:*.ext" string format
    var aepFiles = File.openDialog(
        "Select SynthEyes .aep files to import",
        function(f) { return (f instanceof Folder) || /\.aep$/i.test(f.name); },
        true  // multiselect
    );
    if (!aepFiles) { return; }
    if (!(aepFiles instanceof Array)) aepFiles = [aepFiles];
    if (aepFiles.length === 0) { return; }

    // ── Get command IDs once ──────────────────────────────────────────────────
    var cmdCopy  = app.findMenuCommandId("Copy")  || 3;
    var cmdPaste = app.findMenuCommandId("Paste") || 4;

    // ── Process ───────────────────────────────────────────────────────────────
    var processed = 0, errors = 0, errorLog = [];

    app.beginUndoGroup("Import SynthEyes AEPs to Shots");

    for (var f = 0; f < aepFiles.length; f++) {
        var aepFile  = aepFiles[f];
        // decodeURI handles percent-encoded chars in File.name on macOS
        var baseName = decodeURI(aepFile.name); // e.g. "VS_002_comp_stabilized.jsx.aep"
        var shotName = matchShotName(baseName, shotNames);

        if (!shotName) {
            errors++;
            errorLog.push("• " + baseName + "\n  Could not match any shot folder. Skipped.");
            continue;
        }

        // Find the pre-existing shot folder (e.g. Shots/VS_002)
        var shotFolder = findInFolder(shotsFolder, shotName, FolderItem);
        if (!shotFolder) {
            errors++;
            errorLog.push("• " + baseName + "\n  Shot folder \"" + shotName + "\" not found in Shots.");
            continue;
        }

        try {
            // ── 1. Import the .aep ────────────────────────────────────────────
            // Snapshot existing folder IDs so we can identify the newly added one
            var folderIdsBefore = {};
            for (var n = 1; n <= app.project.numItems; n++) {
                if (app.project.item(n) instanceof FolderItem) {
                    folderIdsBefore[app.project.item(n).id] = true;
                }
            }

            var opts = new ImportOptions(aepFile);
            opts.importAs = ImportAsType.PROJECT;
            app.project.importFile(opts);

            // Find the newly added root folder (first FolderItem with a new ID)
            var importedFolder = null;
            for (var n = 1; n <= app.project.numItems; n++) {
                var candidate = app.project.item(n);
                if (candidate instanceof FolderItem && !folderIdsBefore[candidate.id]) {
                    importedFolder = candidate;
                    break;
                }
            }

            if (!importedFolder) {
                errors++;
                errorLog.push("• " + baseName + "\n  Import returned no new folder — file may already be imported.");
                continue;
            }

            // ── 2. Move imported folder into the shot folder ──────────────────
            importedFolder.parentFolder = shotFolder;

            // ── 3. Find Camera01_3D inside the imported folder ────────────────
            var cameraComp = findCompInFolder(importedFolder, CAMERA_COMP_NAME);
            if (!cameraComp) {
                errorLog.push("• " + baseName + "\n  \"" + CAMERA_COMP_NAME
                    + "\" comp not found — folder moved but layers not copied.");
                errors++;
                processed++;
                continue;
            }

            // ── 4. Find {shotName}_comp in the shot folder ────────────────────
            var targetCompName = shotName + TARGET_COMP_SUFFIX;
            var targetComp = findInFolder(shotFolder, targetCompName, CompItem);
            if (!targetComp) {
                errorLog.push("• " + baseName + "\n  Target comp \"" + targetCompName
                    + "\" not found — folder moved but layers not copied.");
                errors++;
                processed++;
                continue;
            }

            // ── 5. Copy Camera01_3D layers → targetComp at shot clip start ───

            // Open Camera01_3D, select all its layers, copy
            cameraComp.openInViewer();
            for (var l = 1; l <= cameraComp.numLayers; l++) {
                cameraComp.layer(l).selected = true;
            }
            app.executeCommand(cmdCopy);

            // Open target comp and paste
            targetComp.openInViewer();
            app.executeCommand(cmdPaste);

            // Pasted layers are still selected — collect them
            var pastedLayers = [];
            for (var pl = 1; pl <= targetComp.numLayers; pl++) {
                if (targetComp.layer(pl).selected) {
                    pastedLayers.push(targetComp.layer(pl));
                }
            }

            // Find the earliest startTime among pasted layers
            if (pastedLayers.length > 0) {
                var earliestStart = pastedLayers[0].startTime;
                for (var pl = 1; pl < pastedLayers.length; pl++) {
                    if (pastedLayers[pl].startTime < earliestStart) {
                        earliestStart = pastedLayers[pl].startTime;
                    }
                }

                // Find the shot clip layer in the target comp —
                // the (non-pasted) layer whose name starts with the shot name
                // (e.g. "VS_002_plate", "VS_002_footage", …)
                var shotClipStart = targetComp.workAreaStart; // fallback
                for (var sl = 1; sl <= targetComp.numLayers; sl++) {
                    var lyr = targetComp.layer(sl);
                    if (!lyr.selected && lyr.name.indexOf(shotName) === 0) {
                        shotClipStart = lyr.startTime;
                        break;
                    }
                }

                // Shift all pasted layers so the earliest one aligns with the shot clip
                var offset = shotClipStart - earliestStart;
                for (var pl = 0; pl < pastedLayers.length; pl++) {
                    pastedLayers[pl].startTime += offset;
                }
            }

            processed++;

        } catch (e) {
            errors++;
            errorLog.push("• " + baseName + "\n  Error: " + e.message);
        }
    }

    app.endUndoGroup();

    // ── Summary ───────────────────────────────────────────────────────────────
    var summary = "Import complete!\n\n"
        + "✓ Processed : " + processed + "\n"
        + "✗ Errors    : " + errors;
    if (errorLog.length > 0) {
        summary += "\n\n" + errorLog.join("\n\n");
    }
    alert(summary);

})();
