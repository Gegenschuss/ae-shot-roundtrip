/*
       _____                          __
      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
             /___/

  IMPORT RENDERS & GRADES
  After Effects ExtendScript
================================================================================

PURPOSE
-------
Scans all *_comp compositions in the AE project and imports finished
returns from two sources:

  1. each shot's render/ folder    – VFX returns from Nuke etc.
  2. a flat {shots}/_grades/ folder – Resolve-graded returns (optional),
                                      matched to shots by filename prefix

Layers are stacked in the comp with a fixed category order:

    Top  →  grade   (from _grades/, matched by filename prefix)
            render  (from {shot}/render/)
            plate   (disabled when anything sits above)

Within each category the newest file wins; older versions are imported
but disabled.  Category order is fixed, so a VFX re-render after a grade
does not cover the grade.

NAMING CONVENTION (STRICT)
--------------------------
This tool relies on the strict _comp suffix to map comps to disk folders:

    {prefix}_010_comp    ->   {shots}/{prefix}_010/
    {prefix}_010_comp_OS ->   {shots}/{prefix}_010/   (overscan variant)

Comps that do not end in _comp (or _comp_OS) are ignored entirely.

Grades in the flat {shots}/_grades/ folder are matched to the correct
comp by filename prefix — a file must begin with the shot name, e.g.
    KM_010_grade_v01.mov  ->  comp "KM_010_comp"

FOLDER STRUCTURE
----------------
The script expects the same layout the Shot Roundtrip creates, plus an
optional flat _grades/ folder for Resolve returns:

    {shots}/              (default: "../Roundtrip")
      {prefix}_010/
        plate/            <- rendered plate (.mov)
        render/           <- VFX return renders go here
      _grades/            <- Resolve graded returns for all shots
        {prefix}_010_grade_v01.mov
        {prefix}_020_grade_v01.mov

================================================================================
*/

(function () {

    var proj = app.project;
    if (!proj || !proj.file) { alert("Import Renders & Grades: save the project first."); return; }

    // ── UI ────────────────────────────────────────────────────────────────────
    var LABEL_W = 120; var FIELD_H = 22;

    var addRow = function (parent, labelText) {
        var g = parent.add("group");
        g.orientation = "row"; g.alignChildren = ["left", "center"]; g.spacing = 8;
        var lbl = g.add("statictext", undefined, labelText);
        lbl.preferredSize = [LABEL_W, FIELD_H];
        return g;
    };

    var dlg = new Window("dialog", "Gegenschuss Import Renders & Grades");
    dlg.orientation = "column"; dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 10; dlg.margins = 14;

    var pnl = dlg.add("panel", undefined, "Settings");
    pnl.orientation = "column"; pnl.alignChildren = ["fill", "top"];
    pnl.spacing = 6; pnl.margins = [10, 15, 10, 10];

    var r1 = addRow(pnl, "Shots Folder:");
    var etShotsFolder = r1.add("edittext", undefined, "../Roundtrip");
    etShotsFolder.preferredSize = [150, FIELD_H];

    var r2 = addRow(pnl, "File Filter:");
    var etFilter = r2.add("edittext", undefined, "*.mov");
    etFilter.preferredSize = [150, FIELD_H];

    var chkDryRun = pnl.add("checkbox", undefined, "Dry run (preview only, no changes)");
    chkDryRun.value = false;

    var chkReenablePlate = pnl.add("checkbox", undefined, "Re-enable plate layers with no renders");
    chkReenablePlate.value = true;

    var btnGrp = dlg.add("group");
    btnGrp.orientation = "row"; btnGrp.alignment = ["fill", "bottom"]; btnGrp.margins = [0, 4, 0, 0];
    var btnCancel = btnGrp.add("button", undefined, "Cancel");         btnCancel.preferredSize = [80, 28];
    var btnSpacer = btnGrp.add("statictext", undefined, "");           btnSpacer.alignment = ["fill", "center"];
    var btnOk     = btnGrp.add("button", undefined, "Import Renders & Grades"); btnOk.preferredSize = [140, 28];

    btnOk.onClick     = function () { dlg.close(1); };
    btnCancel.onClick = function () { dlg.close(2); };

    if (dlg.show() !== 1) return;

    // ── Resolve paths ─────────────────────────────────────────────────────────
    var aepFolder = proj.file.parent;
    var fsShots = new Folder(aepFolder.fsName + "/" + etShotsFolder.text);
    if (!fsShots.exists) { alert("Shots folder not found:\n" + fsShots.fsName); return; }

    var fileFilter     = etFilter.text;
    var dryRun         = chkDryRun.value;
    var reenablePlate  = chkReenablePlate.value;

    // ── Helpers ───────────────────────────────────────────────────────────────

    // Find or create a folder item in the project.
    function getBinFolder(name) {
        for (var i = 1; i <= proj.numItems; i++) {
            if (proj.item(i).name === name && proj.item(i) instanceof FolderItem) return proj.item(i);
        }
        return proj.items.addFolder(name);
    }

    // Find a child folder by name under a parent bin.
    function getChildBin(parentBin, childName) {
        for (var i = 1; i <= proj.numItems; i++) {
            var it = proj.item(i);
            if (it instanceof FolderItem && it.name === childName && it.parentFolder === parentBin) return it;
        }
        return null;
    }

    // Find or create a child folder under a parent bin.
    function getOrCreateChildBin(parentBin, childName) {
        var existing = getChildBin(parentBin, childName);
        if (existing) return existing;
        var f = proj.items.addFolder(childName);
        f.parentFolder = parentBin;
        return f;
    }

    // Collect all _comp and _comp_OS comps from project.
    function collectShotComps() {
        var result = [];
        for (var i = 1; i <= proj.numItems; i++) {
            var item = proj.item(i);
            if (item instanceof CompItem && /_comp(_OS)?$/i.test(item.name)) {
                result.push(item);
            }
        }
        result.sort(function (a, b) { return a.name < b.name ? -1 : (a.name > b.name ? 1 : 0); });
        return result;
    }

    // Derive the shot name from a comp name: strip _comp or _comp_OS suffix.
    function shotNameFromComp(compName) {
        return compName.replace(/_comp(_OS)?$/i, "");
    }

    // Build a map of file paths already imported as footage in the project.
    function buildImportedFilesMap() {
        var map = {};
        for (var i = 1; i <= proj.numItems; i++) {
            var item = proj.item(i);
            if (item instanceof FootageItem && item.file) {
                map[item.file.fsName] = item;
            }
        }
        return map;
    }

    // Check if a footage item is already used as a layer in a comp.
    function isFootageInComp(comp, footageItem) {
        for (var i = 1; i <= comp.numLayers; i++) {
            try {
                if (comp.layer(i).source === footageItem) return true;
            } catch (e) {}
        }
        return false;
    }

    // Find the plate layer in a comp: bottommost footage layer whose source
    // file path contains /plate/ or whose name contains _plate.
    function findPlateLayer(comp) {
        var candidate = null;
        for (var i = 1; i <= comp.numLayers; i++) {
            var L = comp.layer(i);
            if (L.guideLayer) continue;
            try {
                if (L.source instanceof FootageItem && L.source.file) {
                    var fp = L.source.file.fsName;
                    if (/[\/\\]plate[\/\\]/i.test(fp) || /_plate\./i.test(L.source.name)) {
                        candidate = L; // keep scanning — we want the bottommost match
                    }
                }
            } catch (e) {}
        }
        return candidate;
    }

    // Find all layers in a comp whose source path contains the given segment
    // (e.g. "render" for {shot}/render/, "_grades" for {shots}/_grades/).
    // Returns them in top-to-bottom layer order.
    function findLayersByPathSegment(comp, segment) {
        var esc = segment.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        var re = new RegExp("[\\/\\\\]" + esc + "[\\/\\\\]", "i");
        var result = [];
        for (var i = 1; i <= comp.numLayers; i++) {
            var L = comp.layer(i);
            try {
                if (L.source instanceof FootageItem && L.source.file) {
                    if (re.test(L.source.file.fsName)) result.push(L);
                }
            } catch (e) {}
        }
        return result;
    }

    // Any non-plate return layer stacked above the plate (render or grade).
    function findRenderLayers(comp) {
        return findLayersByPathSegment(comp, "render")
                 .concat(findLayersByPathSegment(comp, "_grades"));
    }

    // ── Main ──────────────────────────────────────────────────────────────────
    var shotComps = collectShotComps();
    if (shotComps.length === 0) { alert("No *_comp compositions found in the project."); return; }

    var importedFiles = buildImportedFilesMap();
    var binShots = getBinFolder("Shots");

    var statsImported      = 0;
    var statsSkipped       = 0;
    var statsComps         = 0;
    var statsReenabled     = 0;
    var compsWithNewRenders = [];
    var dryRunLog          = [];
    var errorLog           = [];

    if (!dryRun) app.beginUndoGroup("Import Renders & Grades");
    try {

        // Two return sources per comp:
        //   render → per-shot folder {shots}/{shot}/render/   (all files)
        //   grade  → flat folder {shots}/_grades/            (files whose name
        //                                                     starts with shot)
        //
        // Processed bottom → top. AE's comp.layers.add() always puts the new
        // layer at the top of the stack, so processing "render" before "grade"
        // guarantees grades end up above VFX returns. Within a category, files
        // are sorted oldest-first so the newest ends up topmost within its group.

        // Pre-read the flat grades folder once per run. Safety-net scaffold
        // in case the user runs this before Shot Roundtrip ever created
        // _grades/. See lib/write_readmes.jsx for the README content.
        var gradesDir = new Folder(fsShots.fsName + "/_grades");
        if (!gradesDir.exists) {
            gradesDir.create();
            try {
                var readmeHelper = new File((new File($.fileName)).parent.parent.fsName + "/lib/write_readmes.jsx");
                if (readmeHelper.exists) $.evalFile(readmeHelper);
                if (typeof writeGradesReadme === "function") writeGradesReadme(gradesDir);
            } catch (eG) { /* non-fatal */ }
        }
        var gradesFiles = gradesDir.getFiles(fileFilter) || [];

        // A filename matches a shot when it starts with the shot name AND the
        // next character is a separator (e.g. "_"), not another letter/digit.
        // Prevents "KM_01" from matching grades for "KM_010".
        function nameStartsWithShot(filename, shotName) {
            var fn = filename.toLowerCase();
            var sn = shotName.toLowerCase();
            if (fn.indexOf(sn) !== 0) return false;
            if (fn.length === sn.length) return true;
            return !/[a-z0-9]/.test(fn.charAt(sn.length));
        }

        function filesForCategory(category, shotName) {
            if (category === "render") {
                var renderDir = new Folder(fsShots.fsName + "/" + shotName + "/render");
                if (!renderDir.exists) return [];
                return renderDir.getFiles(fileFilter) || [];
            }
            // grade: filter flat _grades/ by filename prefix (shot name).
            var matched = [];
            for (var i = 0; i < gradesFiles.length; i++) {
                if (nameStartsWithShot(gradesFiles[i].name, shotName)) {
                    matched.push(gradesFiles[i]);
                }
            }
            return matched;
        }

        var CATEGORIES = ["render", "grade"];

        for (var ci = 0; ci < shotComps.length; ci++) {
            var comp     = shotComps[ci];
            var shotName = shotNameFromComp(comp.name);
            var plateLayer = findPlateLayer(comp);
            var shotBin    = getChildBin(binShots, shotName);
            var gradesBin  = null; // created lazily if any grade imports happen
            var addedAny   = false;
            var foundAnyCategoryFiles = false;

            for (var catIdx = 0; catIdx < CATEGORIES.length; catIdx++) {
                var category = CATEGORIES[catIdx];
                var catFiles = filesForCategory(category, shotName);
                if (catFiles.length === 0) continue;

                // Oldest first → newest ends up as topmost layer within category.
                catFiles.sort(function (a, b) { return a.modified.getTime() - b.modified.getTime(); });

                foundAnyCategoryFiles = true;

                // Destination bin in the AE project panel.
                var destBin;
                if (category === "render") {
                    destBin = shotBin ? getOrCreateChildBin(shotBin, "render") : null;
                } else {
                    if (!gradesBin) gradesBin = getOrCreateChildBin(binShots, "_grades");
                    destBin = gradesBin;
                }

                for (var fi = 0; fi < catFiles.length; fi++) {
                    var catFile = catFiles[fi];
                    var fsPath  = catFile.fsName;
                    var footageItem = importedFiles[fsPath] || null;

                    if (footageItem && isFootageInComp(comp, footageItem)) {
                        statsSkipped++;
                        continue;
                    }

                    if (dryRun) {
                        dryRunLog.push(shotName + " [" + category + "]: " + catFile.name
                            + (footageItem ? " (add layer)" : " (import + add)"));
                        statsImported++;
                        continue;
                    }

                    if (!footageItem) {
                        try {
                            footageItem = proj.importFile(new ImportOptions(catFile));
                            if (destBin) footageItem.parentFolder = destBin;
                            else if (shotBin) footageItem.parentFolder = shotBin;
                            footageItem.label = (category === "grade") ? 11 : 9; // purple grade, green render
                            importedFiles[fsPath] = footageItem;
                        } catch (e) {
                            errorLog.push(shotName + " [" + category + "]: failed to import " + catFile.name + " — " + e.message);
                            continue;
                        }
                    }

                    try {
                        var newLayer = comp.layers.add(footageItem);
                        newLayer.startTime = plateLayer ? plateLayer.startTime : 0;
                        newLayer.position.setValue([comp.width / 2, comp.height / 2]);
                        newLayer.label = (category === "grade") ? 11 : 9;
                        if (plateLayer) {
                            var plateMkr = plateLayer.property("Marker");
                            if (plateMkr && plateMkr.numKeys > 0) {
                                var newMkr = newLayer.property("Marker");
                                for (var mi = 1; mi <= plateMkr.numKeys; mi++) {
                                    try { newMkr.setValueAtTime(plateMkr.keyTime(mi), plateMkr.keyValue(mi)); } catch (e) {}
                                }
                            }
                        }
                        addedAny = true;
                        statsImported++;
                    } catch (e) {
                        errorLog.push(shotName + " [" + category + "]: failed to add layer " + catFile.name + " — " + e.message);
                    }
                }
            }

            // If no category folder had any files, handle plate re-enable and skip the rest.
            if (!foundAnyCategoryFiles) {
                if (reenablePlate && !dryRun) {
                    var pl = findPlateLayer(comp);
                    if (pl && !pl.enabled && findRenderLayers(comp).length === 0) {
                        pl.enabled = true;
                        statsReenabled++;
                    }
                }
                continue;
            }

            statsComps++;
            if (dryRun) continue;

            // Within each category, keep only the topmost (newest) layer enabled.
            var CATEGORY_SEGMENTS = { render: "render", grade: "_grades" };
            for (var ck = 0; ck < CATEGORIES.length; ck++) {
                var catLayers = findLayersByPathSegment(comp, CATEGORY_SEGMENTS[CATEGORIES[ck]]);
                for (var li = 1; li < catLayers.length; li++) {
                    catLayers[li].enabled = false;
                    catLayers[li].audioEnabled = false;
                }
            }

            // Disable plate if anything is stacked above it.
            if (plateLayer) {
                var hasRenders = findRenderLayers(comp).length > 0;
                if (hasRenders) {
                    plateLayer.enabled = false;
                    plateLayer.audioEnabled = false;
                } else if (reenablePlate && !plateLayer.enabled) {
                    plateLayer.enabled = true;
                    statsReenabled++;
                }
            }

            // Keep GUIDE_BURNIN on top.
            var bl = comp.layers.byName("GUIDE_BURNIN");
            if (bl) { bl.locked = false; bl.moveToBeginning(); bl.locked = true; }

            if (addedAny) compsWithNewRenders.push(shotName);
        }

    } catch (e) {
        alert("Import Renders & Grades error:\n" + e.message + "\nLine: " + e.line);
    }
    if (!dryRun) app.endUndoGroup();

    // ── Summary ───────────────────────────────────────────────────────────────
    var title = dryRun ? "Import Renders & Grades — Dry Run" : "Import Renders & Grades Complete";

    var msg = "Comps scanned:       " + shotComps.length + "\n"
            + "Comps with renders:  " + statsComps + "\n"
            + "Renders to import:   " + statsImported + "\n"
            + "Already present:     " + statsSkipped;

    if (statsReenabled > 0) {
        msg += "\nPlates re-enabled:   " + statsReenabled;
    }

    if (dryRunLog.length > 0) {
        msg += "\n\nWould import:";
        for (var di = 0; di < dryRunLog.length; di++) msg += "\n  " + dryRunLog[di];
    }

    if (compsWithNewRenders.length > 0) {
        msg += "\n\nNew renders in:";
        for (var ci2 = 0; ci2 < compsWithNewRenders.length; ci2++) msg += "\n  " + compsWithNewRenders[ci2];
    }

    if (errorLog.length > 0) {
        msg += "\n\nErrors:";
        for (var ei = 0; ei < errorLog.length; ei++) msg += "\n  " + errorLog[ei];
    }

    var sumDlg = new Window("dialog", title);
    sumDlg.orientation = "column";
    sumDlg.alignChildren = ["fill", "top"];
    sumDlg.spacing = 10;
    sumDlg.margins = 14;

    var msgLines = msg.split("\n").length;
    var lineH = 18;
    var maxH = $.screens[0].bottom - $.screens[0].top - 200; // leave room for title bar + button
    var textH = Math.min(msgLines * lineH + 10, maxH);

    var sumText = sumDlg.add("edittext", undefined, msg, { multiline: true, readonly: true, scrollable: true });
    sumText.preferredSize = [400, textH];

    var sumBtnGrp = sumDlg.add("group");
    sumBtnGrp.alignment = ["right", "bottom"];
    var sumOk = sumBtnGrp.add("button", undefined, "OK");
    sumOk.onClick = function () { sumDlg.close(); };

    sumDlg.show();

})();
