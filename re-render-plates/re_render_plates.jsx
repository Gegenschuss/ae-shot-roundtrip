/*
       _____                          __
      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
             /___/

  RE-RENDER PLATES
  After Effects ExtendScript
================================================================================

PURPOSE
-------
Re-render the ORIGINAL plate out of every shot comp to a new filename,
ready to be handed to an external tool (Neat Video, Mocha, a stabilizer,
etc.) and brought back in. The suffix (default "denoised") tags the
output:

    {shots}/{shot}/plate/{shot}_{suffix}.mov
    {shots}/{shot}/plate/{shot}_{suffix}_OS.mov   (overscan variant)

The file lives next to the original plate so it travels with the shot
on handoff — Import Renders & Grades' plate-detection picks it up
automatically as a plate variant (path contains /plate/).

WORKFLOW
--------
1. Run this script. It scans every {shot}_comp (and {shot}_comp_OS)
   for a {shot}_plate precomp and shows a Confirm Shots dialog so
   you can pick which ones to re-render. Confirmed plates get
   {shot}_{suffix}.mov written next to them.
2. The rendered file is imported back into the {shot}_plate precomp
   AND stacked above the original plate — below any renders/grades.
3. Take the rendered file to the external tool, process it, and
   overwrite in place (keep the filename).
4. In AE: select the newly imported footage item and Reload Footage
   (File → Reload Footage) — the processed pixels appear instantly in
   every shot that references it, above the original plate but below
   any VFX returns / grades, so the full stack keeps working.

On re-runs, the script always re-renders the ORIGINAL plate (the
bottom-most tagged plate inside the precomp), never a previously
imported variant — so running it twice gives you a fresh raw plate,
not a copy of whatever the last external pass produced.

WHAT GETS RENDERED
------------------
The outer {shot}_comp, over its workArea (the clip + handles range
set by Shot Roundtrip). Every layer except the raw plate is disabled
for the duration of the render (guide layers are already excluded by
AE), so the output is the pristine plate — never a grade, render, or
plate variant. Enabled states are restored via try/finally.

Shot Roundtrip leaves the raw plate flat + locked at the bottom of
`_comp`. The `{shot}_plate` precomp is created lazily — this script
creates it at import time (first reimport lands inside it) so the
round-tripped variant stacks in the same container as grades and VFX
renders.

VERSIONING
----------
Because we mutate the project (precompose + add layers), the script
Save-As's the .aep to the next _v## before any DOM change. Your
original is preserved as the rollback point. Dry-run skips the bump
and runs in detect-only mode.

================================================================================
*/

(function () {

    var proj = app.project;
    if (!proj || !proj.file) { alert("Re-render Plates: save the project first."); return; }

    // ── UI ────────────────────────────────────────────────────────────────────
    var LABEL_W = 120; var FIELD_H = 22;

    var addRow = function (parent, labelText) {
        var g = parent.add("group");
        g.orientation = "row"; g.alignChildren = ["left", "center"]; g.spacing = 8;
        var lbl = g.add("statictext", undefined, labelText);
        lbl.preferredSize = [LABEL_W, FIELD_H];
        return g;
    };

    var dlg = new Window("dialog", "Gegenschuss Re-render Plates");
    dlg.orientation = "column"; dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 10; dlg.margins = 14;

    var pnl = dlg.add("panel", undefined, "Settings");
    pnl.orientation = "column"; pnl.alignChildren = ["fill", "top"];
    pnl.spacing = 6; pnl.margins = [10, 15, 10, 10];

    var r1 = addRow(pnl, "Suffix:");
    var etSuffix = r1.add("edittext", undefined, "denoised");
    etSuffix.preferredSize = [150, FIELD_H];
    var suffixHint = r1.add("statictext", undefined, "\u2192 {shot}_{suffix}.mov");
    suffixHint.graphics.foregroundColor = suffixHint.graphics.newPen(
        suffixHint.graphics.PenType.SOLID_COLOR, [0.55, 0.55, 0.55, 1], 1);

    var r2 = addRow(pnl, "Shots Folder:");
    var etShotsFolder = r2.add("edittext", undefined, "../Roundtrip");
    etShotsFolder.preferredSize = [150, FIELD_H];
    var btnBrowseShots = r2.add("button", undefined, "Browse\u2026");
    btnBrowseShots.preferredSize = [80, FIELD_H];
    btnBrowseShots.onClick = function () {
        var seed = null;
        try {
            var txt = etShotsFolder.text || "";
            var candidate = /^(\/|[A-Za-z]:)/.test(txt)
                          ? new Folder(txt)
                          : (proj && proj.file ? new Folder(proj.file.parent.fsName + "/" + txt) : null);
            if (candidate && candidate.exists) seed = candidate;
        } catch (eSeed) {}
        var picked = (seed ? seed : Folder.desktop).selectDlg("Select Shots Folder");
        if (picked) etShotsFolder.text = picked.fsName;
    };

    var r3 = addRow(pnl, "OM Template:");
    var etOM = r3.add("edittext", undefined, "ProRes 422 HQ");
    etOM.preferredSize = [150, FIELD_H];

    var chkDryRun = pnl.add("checkbox", undefined, "Dry run (preview only, no changes)");
    chkDryRun.value = false;

    // ── Settings persistence ──────────────────────────────
    // Suffix, Shots Folder, and OM Template round-trip through app.settings.
    // Dry Run (debug) isn't persisted. Reset restores defaults in the fields
    // without saving until the user clicks Re-render.
    var RR_SECTION  = "Gegenschuss Re-render Plates";
    var RR_DEFAULTS = {
        suffix:      "denoised",
        shotsFolder: "../Roundtrip",
        omTemplate:  "ProRes 422 HQ"
    };
    function rrLoad(key, fallback) {
        try {
            if (app.settings.haveSetting(RR_SECTION, key)) return app.settings.getSetting(RR_SECTION, key);
        } catch (e) {}
        return fallback;
    }
    function rrSave(key, value) {
        try { app.settings.saveSetting(RR_SECTION, key, String(value)); } catch (e) {}
    }
    function rrApply(s) {
        etSuffix.text      = s.suffix;
        etShotsFolder.text = s.shotsFolder;
        etOM.text          = s.omTemplate;
    }
    rrApply({
        suffix:      rrLoad("suffix",      RR_DEFAULTS.suffix),
        shotsFolder: rrLoad("shotsFolder", RR_DEFAULTS.shotsFolder),
        omTemplate:  rrLoad("omTemplate",  RR_DEFAULTS.omTemplate)
    });

    var btnGrp = dlg.add("group");
    btnGrp.orientation = "row"; btnGrp.alignment = ["fill", "bottom"]; btnGrp.margins = [0, 4, 0, 0];
    var btnReset  = btnGrp.add("button", undefined, "Reset to Defaults"); btnReset.preferredSize  = [130, 28];
    var btnCancel = btnGrp.add("button", undefined, "Cancel");            btnCancel.preferredSize = [80, 28];
    var btnSpacer = btnGrp.add("statictext", undefined, "");              btnSpacer.alignment = ["fill", "center"];
    var btnOk     = btnGrp.add("button", undefined, "Re-render Plates");  btnOk.preferredSize = [140, 28];

    btnReset.onClick  = function () { rrApply(RR_DEFAULTS); };
    btnOk.onClick     = function () {
        rrSave("suffix",      etSuffix.text);
        rrSave("shotsFolder", etShotsFolder.text);
        rrSave("omTemplate",  etOM.text);
        dlg.close(1);
    };
    btnCancel.onClick = function () { dlg.close(2); };

    if (dlg.show() !== 1) return;

    // ── Resolve inputs ────────────────────────────────────────────────────────
    var suffix = String(etSuffix.text || "").replace(/^[\s_]+|[\s_]+$/g, "");
    if (!suffix) { alert("Re-render Plates: suffix is required."); return; }
    // Sanitise for a filename segment — strip anything a filesystem hates and
    // collapse internal whitespace to underscores.
    suffix = suffix.replace(/\s+/g, "_").replace(/[\/\\:*?"<>|]/g, "");

    var aepFolder = proj.file.parent;
    var shotsPathText = etShotsFolder.text;
    var fsShots = /^(\/|[A-Za-z]:)/.test(shotsPathText)
                ? new Folder(shotsPathText)
                : new Folder(aepFolder.fsName + "/" + shotsPathText);
    if (!fsShots.exists) { alert("Shots folder not found:\n" + fsShots.fsName); return; }

    var omTemplate = etOM.text;
    var dryRun     = chkDryRun.value;

    // ── Project helpers ───────────────────────────────────────────────────────

    function getBinFolder(name) {
        for (var i = 1; i <= proj.numItems; i++) {
            if (proj.item(i).name === name && proj.item(i) instanceof FolderItem) return proj.item(i);
        }
        return proj.items.addFolder(name);
    }

    function getChildBin(parentBin, childName) {
        for (var i = 1; i <= proj.numItems; i++) {
            var it = proj.item(i);
            if (it instanceof FolderItem && it.name === childName && it.parentFolder === parentBin) return it;
        }
        return null;
    }

    function getOrCreateChildBin(parentBin, childName) {
        var existing = getChildBin(parentBin, childName);
        if (existing) return existing;
        var f = proj.items.addFolder(childName);
        f.parentFolder = parentBin;
        return f;
    }

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

    function shotNameFromComp(compName) {
        return compName.replace(/_comp(_OS)?$/i, "");
    }

    function isOvercan(compName) {
        return /_comp_OS$/i.test(compName);
    }

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

    function isFootageInComp(comp, footageItem) {
        for (var i = 1; i <= comp.numLayers; i++) {
            try {
                if (comp.layer(i).source === footageItem) return true;
            } catch (e) {}
        }
        return false;
    }

    // Mirror of import_renders.jsx findPlateLayer: four-tier fallback to
    // locate the SOURCE plate — the one we want to render out, never a
    // previously imported variant. Tier 1 (path contains /plate/ or filename
    // ends _plate.) is what Shot Roundtrip's own plates match, so re-runs
    // always hit the pristine original.
    function findPlateLayer(comp, shotName) {
        var shotPrefix = (shotName || "").toLowerCase();
        var tagged = null, shotMatch = null, untagged = null, anyFootage = null;
        for (var i = 1; i <= comp.numLayers; i++) {
            var L = comp.layer(i);
            if (L.guideLayer) continue;
            try {
                if (!(L.source instanceof FootageItem) || !L.source.file) continue;
                var fp = L.source.file.fsName;
                anyFootage = L;
                if (/[\/\\]plate[\/\\]/i.test(fp) || /_plate\./i.test(L.source.name)) {
                    tagged = L;
                    continue;
                }
                if (/[\/\\]render[\/\\]/i.test(fp) || /[\/\\]_grade[\/\\]/i.test(fp)) continue;
                if (shotPrefix && L.source.name &&
                    L.source.name.toLowerCase().indexOf(shotPrefix) === 0) {
                    shotMatch = L;
                    continue;
                }
                untagged = L;
            } catch (e) {}
        }
        return tagged || shotMatch || untagged || anyFootage;
    }

    // Topmost plate-like layer (non-render, non-grade) — anchor for where
    // newly imported plate variants stack. Mirrors import_renders.jsx.
    function findTopmostPlateLike(comp) {
        for (var i = 1; i <= comp.numLayers; i++) {
            var L = comp.layer(i);
            if (L.guideLayer) continue;
            try {
                if (!(L.source instanceof FootageItem) || !L.source.file) continue;
                var fp = L.source.file.fsName;
                if (/[\/\\]render[\/\\]/i.test(fp) || /[\/\\]_grade[\/\\]/i.test(fp)) continue;
                return L;
            } catch (e) {}
        }
        return null;
    }

    // Derive the plate-precomp name from a _comp name.
    function plateCompNameFor(compName) {
        return compName.replace(/_comp(_OS)?$/i, function (_m, os) {
            return "_plate" + (os || "");
        });
    }

    // Ensure a {shot}_plate precomp exists as a layer at the TOP of `comp`.
    // Empty on first creation — holds only reimported variants (grades, VFX,
    // denoised plates). Raw plate stays flat + locked at the bottom of `comp`.
    // Mirror of import_renders.jsx's ensurePlatePrecomp so Re-render Plates can
    // land its output in the same place. readOnly returns null if missing.
    function ensurePlatePrecomp(comp, shotName, readOnly) {
        var targetName = plateCompNameFor(comp.name);
        for (var i = 1; i <= comp.numLayers; i++) {
            var L = comp.layer(i);
            try {
                if (L.source instanceof CompItem && L.source.name === targetName) return L.source;
            } catch (e) {}
        }
        if (readOnly) return null;
        var newComp = proj.items.addComp(
            targetName, comp.width, comp.height, comp.pixelAspect,
            comp.duration, comp.frameRate
        );
        newComp.displayStartTime = 0;
        try {
            var shotBin      = getChildBin(binShots, shotName);
            var shotPlateBin = shotBin ? getOrCreateChildBin(shotBin, "plate") : null;
            if (shotPlateBin) newComp.parentFolder = shotPlateBin;
        } catch (eBin) {}
        try {
            if (comp.markerProperty && comp.markerProperty.numKeys > 0) {
                var dstMkr = newComp.markerProperty;
                for (var mi = 1; mi <= comp.markerProperty.numKeys; mi++) {
                    try {
                        dstMkr.setValueAtTime(
                            comp.markerProperty.keyTime(mi),
                            comp.markerProperty.keyValue(mi)
                        );
                    } catch (e) {}
                }
            }
        } catch (eCM) {}
        var newLayer = comp.layers.add(newComp);
        try { newLayer.moveToBeginning(); } catch (eMB) {}
        try { newLayer.startTime = 0; } catch (eST) {}
        return newComp;
    }

    // Confirm Shots preflight: lets the user toggle which shots to re-render.
    // Defaults to all selected (opt-out, since the common case is "re-render
    // everything"). Returns the filtered plan, or null if the user cancelled.
    // Mirrors shot_roundtrip.jsx's Confirm Shots pattern — listbox + Toggle
    // button + Space keystroke.
    function showConfirmDialog(plan) {
        var win = new Window("dialog", "Re-render Plates \u2014 Confirm");
        win.orientation = "column"; win.alignChildren = ["fill", "top"];
        win.spacing = 8; win.margins = 14;

        win.add("statictext", undefined,
            plan.length + " plate" + (plan.length === 1 ? "" : "s")
            + " ready. Select rows and press Space (or Toggle Selected) to toggle "
            + "whether they render. Checked rows render.");

        // Column widths auto-sized from content.
        var shotMax = 0, plateMax = 0, outMax = 0;
        for (var ml = 0; ml < plan.length; ml++) {
            var sn = plan[ml].shotName + (plan[ml].isOS ? " [OS]" : "");
            if (sn.length > shotMax) shotMax = sn.length;
            var pn = (plan[ml].plateLayer.source && plan[ml].plateLayer.source.name) ? plan[ml].plateLayer.source.name : "?";
            if (pn.length > plateMax) plateMax = pn.length;
            var on = plan[ml].outFile.name;
            if (on.length > outMax) outMax = on.length;
        }
        var shotW  = Math.max(Math.min(shotMax  * 8 + 24, 200),  90);
        var plateW = Math.max(Math.min(plateMax * 7 + 24, 340), 160);
        var outW   = Math.max(Math.min(outMax   * 7 + 24, 400), 180);
        var checkW = 60, framesW = 70;

        var lb = win.add("listbox", undefined, [], {
            multiselect: true,
            numberOfColumns: 5,
            showHeaders: true,
            columnTitles: ["Render", "Shot", "Plate", "Frames", "Output"],
            columnWidths: [checkW, shotW, plateW, framesW, outW]
        });
        var lbH = Math.min(Math.max(plan.length * 22 + 40, 180), 520);
        lb.preferredSize = [checkW + shotW + plateW + framesW + outW + 40, lbH];

        var selectedState = [];
        for (var pi = 0; pi < plan.length; pi++) {
            selectedState.push(true);
            var item = lb.add("item", "\u2713");
            item.subItems[0].text = plan[pi].shotName + (plan[pi].isOS ? " [OS]" : "");
            item.subItems[1].text = (plan[pi].plateLayer.source && plan[pi].plateLayer.source.name) ? plan[pi].plateLayer.source.name : "?";
            // Render span = _comp's workArea (clip + handles, set by Shot Roundtrip).
            var fr = plan[pi].comp.frameRate;
            var waDur = 0; try { waDur = plan[pi].comp.workAreaDuration; } catch (eWA) {}
            var frames = (fr > 0 && waDur > 0) ? Math.round(waDur * fr) : 0;
            item.subItems[2].text = String(frames);
            item.subItems[3].text = plan[pi].outFile.name;
        }

        function toggleSelected() {
            var sel = lb.selection;
            if (!sel) return;
            var arr = (sel.length !== undefined) ? sel : [sel];
            for (var si = 0; si < arr.length; si++) {
                var idx = arr[si].index;
                selectedState[idx] = !selectedState[idx];
                lb.items[idx].text = selectedState[idx] ? "\u2713" : "";
            }
        }

        var btnGrp = win.add("group");
        btnGrp.orientation = "row"; btnGrp.alignment = ["fill", "bottom"];
        var spacer = btnGrp.add("statictext", undefined, ""); spacer.alignment = ["fill", "center"];
        var btnToggle = btnGrp.add("button", undefined, "Toggle Selected"); btnToggle.preferredSize = [130, 28];
        var btnCancel = btnGrp.add("button", undefined, "Cancel");          btnCancel.preferredSize = [80, 28];
        var btnOk     = btnGrp.add("button", undefined, "Re-render");       btnOk.preferredSize     = [110, 28];

        btnToggle.onClick = toggleSelected;
        btnCancel.onClick = function () { win.close(2); };
        btnOk.onClick     = function () { win.close(1); };

        win.addEventListener("keydown", function (e) {
            if (e.keyName === "Space") { e.preventDefault(); toggleSelected(); }
        });

        if (win.show() !== 1) return null;

        var filtered = [];
        for (var fi = 0; fi < plan.length; fi++) {
            if (selectedState[fi]) filtered.push(plan[fi]);
        }
        return filtered;
    }

    // Save As → next _v## version (same pattern as import_renders.jsx).
    function pad(n, s) { var str = "" + n; while (str.length < s) str = "0" + str; return str; }
    function saveAsNextVersion() {
        var cur = proj.file;
        if (!cur) return null;
        var baseName = cur.name.replace(/\.aep$/i, "");
        var m = baseName.match(/^(.*?)(_?v)(\d+)$/);
        var stem, prefix, width, next;
        if (m) {
            stem = m[1]; prefix = m[2]; width = m[3].length;
            next = parseInt(m[3], 10) + 1;
        } else {
            stem = baseName; prefix = "_v"; width = 2; next = 1;
        }
        var newFile = null;
        while (next < 10000) {
            var candidate = new File(cur.parent.fsName + "/" + stem + prefix + pad(next, width) + ".aep");
            if (!candidate.exists) { newFile = candidate; break; }
            next++;
        }
        if (!newFile) {
            alert("Re-render Plates: could not find an unused version number for the backup copy.\nAborting so nothing is modified.");
            return null;
        }
        try { proj.save(); proj.save(newFile); }
        catch (eSave) {
            alert("Re-render Plates: failed to save versioned copy —\n" + eSave.message +
                  "\n\nAborting so the original file stays untouched.");
            return null;
        }
        return newFile;
    }

    function waitForFile(file, tries) {
        for (var i = 0; i < tries; i++) {
            if (file.exists && file.length > 0) return true;
            $.sleep(500);
        }
        return false;
    }

    // ── Main ──────────────────────────────────────────────────────────────────
    var shotComps = collectShotComps();
    if (shotComps.length === 0) { alert("No *_comp compositions found in the project."); return; }

    var binShots = getBinFolder("Shots");

    // Scan to build renderPlan. No project mutations here so the user can
    // still cancel at the Confirm dialog below with nothing touched.
    var dryRunLog  = [];
    var errorLog   = [];
    var renderPlan = [];  // { comp, platePrecomp, plateLayer, outFile, shotName, isOS }

    try {
        for (var ci = 0; ci < shotComps.length; ci++) {
            var comp     = shotComps[ci];
            var shotName = shotNameFromComp(comp.name);
            var isOS     = isOvercan(comp.name);

            // Raw plate is flat + locked in _comp (Shot Roundtrip layout). Find
            // it directly in _comp — the {shot}_plate precomp, if it exists,
            // holds only reimported variants and must not be our render source.
            var srcPlate = findPlateLayer(comp, shotName);
            if (!srcPlate) {
                errorLog.push(shotName + (isOS ? " [OS]" : "")
                    + ": no plate layer found in " + comp.name + "; skipped.");
                continue;
            }

            var outName = shotName + "_" + suffix + (isOS ? "_OS" : "") + ".mov";
            var outFile = new File(fsShots.fsName + "/" + shotName + "/plate/" + outName);

            if (dryRun) {
                dryRunLog.push(shotName + (isOS ? " [OS]" : "") + ": "
                    + "would render " + comp.name + " \u2192 " + outFile.fsName
                    + " (plate layer: " + srcPlate.source.name + ")");
                continue;
            }

            renderPlan.push({
                comp:       comp,
                plateLayer: srcPlate,
                outFile:    outFile,
                shotName:   shotName,
                isOS:       isOS
            });
        }
    } catch (ePrep) {
        errorLog.push("PREP: " + ePrep.message + " (line " + ePrep.line + ")");
    }

    // ── Dry-run exit ──────────────────────────────────────────────────────────
    if (dryRun) {
        var dMsg = "Dry run — no changes made.\n\n"
                 + "Comps scanned:        " + shotComps.length + "\n"
                 + "Plates to render:     " + dryRunLog.length;
        if (dryRunLog.length > 0) {
            dMsg += "\n\nWould render:";
            for (var di = 0; di < dryRunLog.length; di++) dMsg += "\n  " + dryRunLog[di];
        }
        if (errorLog.length > 0) {
            dMsg += "\n\nWarnings:";
            for (var ei = 0; ei < errorLog.length; ei++) dMsg += "\n  " + errorLog[ei];
        }
        showSummary("Re-render Plates — Dry Run", dMsg);
        return;
    }

    if (renderPlan.length === 0) {
        showSummary("Re-render Plates", "No plates to render.\n\n"
            + (errorLog.length > 0 ? ("Warnings:\n  " + errorLog.join("\n  ")) : ""));
        return;
    }

    // ── Confirm Shots ────────────────────────────────────────────────────────
    // User picks which plates to re-render. Cancel aborts cleanly — nothing
    // has been mutated yet (no version bump, no layer disabling).
    var confirmed = showConfirmDialog(renderPlan);
    if (confirmed === null) return;
    renderPlan = confirmed;
    if (renderPlan.length === 0) {
        showSummary("Re-render Plates", "No shots selected \u2014 nothing to render.");
        return;
    }

    // ── Version bump (AFTER confirm, so cancelled runs don't leave a copy) ───
    var versionedFile = saveAsNextVersion();
    if (!versionedFile) return;

    // ── Isolate plates + ensure output dirs ──────────────────────────────────
    // Snapshot enabled-states inside each _comp so we can restore after the
    // render. Temporarily disable every layer except the raw plate so the
    // render output is the pristine plate — never a grade, render, or plate
    // variant that happens to live in the {shot}_plate precomp above it.
    // Guide layers (GUIDE_BURNIN, the DO-NOT-EDIT note) are auto-excluded.
    app.beginUndoGroup("Re-render Plates: isolate");
    try {
        for (var pi = 0; pi < renderPlan.length; pi++) {
            var rp = renderPlan[pi];

            var shotDir = new Folder(fsShots.fsName + "/" + rp.shotName);
            if (!shotDir.exists) shotDir.create();
            var plateDir = new Folder(shotDir.fsName + "/plate");
            if (!plateDir.exists) plateDir.create();

            var restoreStates = [];
            for (var li = 1; li <= rp.comp.numLayers; li++) {
                var L = rp.comp.layer(li);
                restoreStates.push({
                    layer:   L,
                    enabled: L.enabled,
                    audio:   L.audioEnabled
                });
                if (L !== rp.plateLayer && !L.guideLayer) {
                    try { L.enabled = false; } catch (eEn) {}
                    try { L.audioEnabled = false; } catch (eAu) {}
                }
            }
            // Raw plate is locked; enabled toggles still apply to locked layers.
            try { rp.plateLayer.enabled = true; } catch (ePE) {}
            rp.restoreStates = restoreStates;
        }
    } finally {
        app.endUndoGroup();
    }

    // ── Queue renders ─────────────────────────────────────────────────────────
    var queueErrors = [];
    for (var qi = 0; qi < renderPlan.length; qi++) {
        var entry = renderPlan[qi];
        try {
            // Render the outer _comp over its workArea (= clip + handles set
            // by Shot Roundtrip). Everything-but-plate is disabled above, so
            // the output is the pristine raw plate.
            var rq = proj.renderQueue.items.add(entry.comp);
            rq.timeSpanStart    = entry.comp.workAreaStart;
            rq.timeSpanDuration = entry.comp.workAreaDuration;
            var om = rq.outputModule(1);
            var foundT = false;
            for (var t = 0; t < om.templates.length; t++) if (om.templates[t] === omTemplate) foundT = true;
            if (foundT) om.applyTemplate(omTemplate);
            om.file = entry.outFile;
            entry.rqItem = rq;
        } catch (eQ) {
            queueErrors.push(entry.shotName + (entry.isOS ? " [OS]" : "") + ": queue failed — " + eQ.message);
            entry.queueFailed = true;
        }
    }

    // ── Render ────────────────────────────────────────────────────────────────
    // AE blocks during renderQueue.render(); the render window takes over.
    var renderSucceeded = false;
    try {
        proj.save();
        proj.renderQueue.render();
        renderSucceeded = true;
    } catch (eR) {
        errorLog.push("RENDER: " + eR.message + " — check OM Template and disk space.");
    }

    // ── Restore layer states (always, even on render failure) ─────────────────
    app.beginUndoGroup("Re-render Plates: restore + import");
    try {
        for (var rI = 0; rI < renderPlan.length; rI++) {
            var e = renderPlan[rI];
            for (var rs = 0; rs < e.restoreStates.length; rs++) {
                var st = e.restoreStates[rs];
                try { st.layer.enabled      = st.enabled; } catch (eR1) {}
                try { st.layer.audioEnabled = st.audio;   } catch (eR2) {}
            }
        }
    } catch (eRestore) {
        errorLog.push("RESTORE: " + eRestore.message);
    }

    // ── Import rendered plates back into the _plate precomp ───────────────────
    var imported = 0;
    var reloaded = 0;
    var skipped  = 0;

    if (renderSucceeded) {
        var importedFiles = buildImportedFilesMap();
        for (var iI = 0; iI < renderPlan.length; iI++) {
            var ent = renderPlan[iI];
            if (ent.queueFailed) continue;
            if (!waitForFile(ent.outFile, 20)) {
                errorLog.push(ent.shotName + (ent.isOS ? " [OS]" : "")
                    + ": rendered file not found at " + ent.outFile.fsName);
                continue;
            }

            var fsPath = ent.outFile.fsName;
            var footageItem = importedFiles[fsPath] || null;
            var shotBin = getChildBin(binShots, ent.shotName);
            var plateBin = shotBin ? getOrCreateChildBin(shotBin, "plate") : null;

            if (!footageItem) {
                try {
                    footageItem = proj.importFile(new ImportOptions(ent.outFile));
                    if (plateBin) footageItem.parentFolder = plateBin;
                    footageItem.label = 12; // Sandstone — plate variant (neither grade nor render)
                    importedFiles[fsPath] = footageItem;
                    imported++;
                } catch (eIm) {
                    errorLog.push(ent.shotName + (ent.isOS ? " [OS]" : "")
                        + ": import failed — " + eIm.message);
                    continue;
                }
            } else {
                // File may have been overwritten by a prior run — refresh source.
                try { footageItem.replace(ent.outFile); reloaded++; } catch (eRl) {}
            }

            // Ensure the {shot}_plate precomp exists in _comp (create empty if
            // this is the first reimport for the shot), then add the rendered
            // variant inside it.
            var platePrecomp = ensurePlatePrecomp(ent.comp, ent.shotName, false);
            if (!platePrecomp) {
                errorLog.push(ent.shotName + (ent.isOS ? " [OS]" : "")
                    + ": could not create " + plateCompNameFor(ent.comp.name));
                continue;
            }
            if (isFootageInComp(platePrecomp, footageItem)) {
                skipped++;
                continue;
            }
            try {
                var newLayer = platePrecomp.layers.add(footageItem);
                // Render file duration = _comp.workAreaDuration (clip + handles).
                // Drop it at workAreaStart so source-frame at handle-in lines up
                // with comp time workAreaStart — matches import_renders.jsx's
                // "clip+handles" alignment case.
                try { newLayer.startTime = ent.comp.workAreaStart; } catch (eST) {}
                newLayer.position.setValue([platePrecomp.width / 2, platePrecomp.height / 2]);
                newLayer.label = 12;
                // Plate variants sit above existing plate-like layers but
                // below any renders/grades — matches import_renders.jsx's
                // stack-above-topmost-plate-like convention.
                var plateTop = findTopmostPlateLike(platePrecomp);
                if (plateTop && plateTop !== newLayer) {
                    try { newLayer.moveBefore(plateTop); } catch (eMv) {}
                }
            } catch (eAdd) {
                errorLog.push(ent.shotName + (ent.isOS ? " [OS]" : "")
                    + ": failed to add layer — " + eAdd.message);
            }
        }
    }
    app.endUndoGroup();

    // ── Summary ───────────────────────────────────────────────────────────────
    var queued    = renderPlan.length - queueErrors.length;
    var msg = "Comps scanned:       " + shotComps.length + "\n"
            + "Plates queued:       " + queued + "\n"
            + "Rendered + imported: " + imported + "\n"
            + "Reloaded (overwrite):" + reloaded + "\n"
            + "Already in precomp:  " + skipped;
    if (versionedFile) {
        msg += "\n\nWorking in:          " + versionedFile.name
             + "\n(original preserved on disk as the rollback point)";
    }
    if (queueErrors.length > 0) {
        msg += "\n\nQueue errors:";
        for (var qe = 0; qe < queueErrors.length; qe++) msg += "\n  " + queueErrors[qe];
    }
    if (errorLog.length > 0) {
        msg += "\n\nErrors / warnings:";
        for (var er = 0; er < errorLog.length; er++) msg += "\n  " + errorLog[er];
    }

    showSummary("Re-render Plates Complete", msg);

    function showSummary(title, body) {
        var sumDlg = new Window("dialog", title);
        sumDlg.orientation = "column";
        sumDlg.alignChildren = ["fill", "top"];
        sumDlg.spacing = 10;
        sumDlg.margins = 14;

        var lines = body.split("\n").length;
        var maxH = $.screens[0].bottom - $.screens[0].top - 200;
        var textH = Math.min(lines * 18 + 10, maxH);

        var txt = sumDlg.add("edittext", undefined, body,
            { multiline: true, readonly: true, scrollable: true });
        txt.preferredSize = [460, textH];

        var row = sumDlg.add("group");
        row.alignment = ["right", "bottom"];
        var ok = row.add("button", undefined, "OK");
        ok.onClick = function () { sumDlg.close(); };
        sumDlg.show();
    }

})();
