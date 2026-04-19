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
  2. a flat {shots}/_grade/ folder  – Resolve-graded returns (optional),
                                      matched to shots by filename prefix

Layers are stacked in the comp with a fixed category order:

    Top  →  grade   (from _grade/, matched by filename prefix)
            render  (from {shot}/render/)
            plate   (enabled state is never touched — you manage it)

Within each category the newest file wins; older versions are imported
but disabled.  Category order is fixed, so a VFX re-render after a grade
does not cover the grade. Plate variants (stabilized, denoised, retimed
versions rendered alongside the original) are treated as plates too —
renders and grades always stack above the topmost plate-like layer.

NAMING CONVENTION (STRICT)
--------------------------
This tool relies on the strict _comp suffix to map comps to disk folders:

    {prefix}_010_comp    ->   {shots}/{prefix}_010/
    {prefix}_010_comp_OS ->   {shots}/{prefix}_010/   (overscan variant)

Comps that do not end in _comp (or _comp_OS) are ignored entirely.

Grades in the flat {shots}/_grade/ folder are matched to the correct
comp by filename prefix — a file must begin with the shot name, e.g.
    KM_010_grade_v01.mov  ->  comp "KM_010_comp"

FOLDER STRUCTURE
----------------
The script expects the same layout the Shot Roundtrip creates, plus an
optional flat _grade/ folder for Resolve returns:

    {shots}/              (default: "../Roundtrip")
      {prefix}_010/
        plate/            <- rendered plate (.mov)
        render/           <- VFX return renders go here
      _grade/             <- Resolve graded returns for all shots
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
    var btnBrowseShots = r1.add("button", undefined, "Browse\u2026");
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

    var r2 = addRow(pnl, "File Filter:");
    var etFilter = r2.add("edittext", undefined, "*.mov");
    etFilter.preferredSize = [150, FIELD_H];

    var chkDryRun = pnl.add("checkbox", undefined, "Dry run (preview only, no changes)");
    chkDryRun.value = false;

    // ── Settings persistence ──────────────────────────────
    // Shots Folder and File Filter get round-tripped through app.settings.
    // Dry Run (debug) isn't persisted. Reset restores defaults in the fields
    // without saving until the user clicks Import.
    var IR_SECTION  = "Gegenschuss Import Renders";
    var IR_DEFAULTS = {
        shotsFolder:    "../Roundtrip",
        fileFilter:     "*.mov"
    };
    function irLoad(key, fallback) {
        try {
            if (app.settings.haveSetting(IR_SECTION, key)) return app.settings.getSetting(IR_SECTION, key);
        } catch (e) {}
        return fallback;
    }
    function irSave(key, value) {
        try { app.settings.saveSetting(IR_SECTION, key, String(value)); } catch (e) {}
    }
    function irApply(s) {
        etShotsFolder.text     = s.shotsFolder;
        etFilter.text          = s.fileFilter;
    }
    irApply({
        shotsFolder:   irLoad("shotsFolder",   IR_DEFAULTS.shotsFolder),
        fileFilter:    irLoad("fileFilter",    IR_DEFAULTS.fileFilter)
    });

    var btnGrp = dlg.add("group");
    btnGrp.orientation = "row"; btnGrp.alignment = ["fill", "bottom"]; btnGrp.margins = [0, 4, 0, 0];
    var btnReset  = btnGrp.add("button", undefined, "Reset to Defaults"); btnReset.preferredSize  = [130, 28];
    var btnCancel = btnGrp.add("button", undefined, "Cancel");            btnCancel.preferredSize = [80, 28];
    var btnSpacer = btnGrp.add("statictext", undefined, "");              btnSpacer.alignment = ["fill", "center"];
    var btnOk     = btnGrp.add("button", undefined, "Import Renders & Grades"); btnOk.preferredSize = [140, 28];

    btnReset.onClick  = function () { irApply(IR_DEFAULTS); };
    btnOk.onClick     = function () {
        irSave("shotsFolder",   etShotsFolder.text);
        irSave("fileFilter",    etFilter.text);
        dlg.close(1);
    };
    btnCancel.onClick = function () { dlg.close(2); };

    if (dlg.show() !== 1) return;

    // ── Resolve paths ─────────────────────────────────────────────────────────
    var aepFolder = proj.file.parent;
    // Accept either an absolute path (from Browse) or one relative to the
    // .aep's parent (legacy default "../Roundtrip").
    var shotsPathText = etShotsFolder.text;
    var fsShots = /^(\/|[A-Za-z]:)/.test(shotsPathText)
                ? new Folder(shotsPathText)
                : new Folder(aepFolder.fsName + "/" + shotsPathText);
    if (!fsShots.exists) { alert("Shots folder not found:\n" + fsShots.fsName); return; }

    var fileFilter     = etFilter.text;
    var dryRun         = chkDryRun.value;

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
    // Find the plate layer — the reference layer that new renders/grades will
    // be aligned to (startTime + markers copied from it). Four-tier fallback:
    //   1. Tagged plate: path contains /plate/ or filename has _plate.
    //      (the roundtrip's own convention)
    //   2. Non-render/grade footage whose name STARTS WITH THE SHOT NAME
    //      (e.g. plate named "IP_010.mov" in shot "IP_010"). Prevents
    //      unrelated footage that happens to live in the comp (other plates,
    //      reference clips, etc.) from being adopted.
    //   3. Any non-render/grade footage — plate with an unrelated name.
    //   4. Any footage layer at all — last resort so imports still land on
    //      something sensible instead of comp time 0.
    // Each tier picks the BOTTOMMOST match so stacks with plate at the
    // bottom (VFX convention) work without surprises.
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
                // Exclude renders and grades from the untagged-plate fallbacks;
                // those are outputs, not the source plate reference.
                if (/[\/\\]render[\/\\]/i.test(fp) || /[\/\\]_grade[\/\\]/i.test(fp)) continue;
                // Shot-name prefix match (tier 2) — the most likely plate when
                // there are no /plate/ path or _plate. filename tags.
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

    // Return the topmost existing render-or-grade layer (the current "head"
    // of the VFX/grade stack). New grades stack above this anchor; new
    // renders stack above the topmost plate-like layer directly.
    function findTopmostRenderOrGrade(comp) {
        for (var i = 1; i <= comp.numLayers; i++) {
            var L = comp.layer(i);
            if (L.guideLayer) continue;
            try {
                if (L.source instanceof FootageItem && L.source.file) {
                    var fp = L.source.file.fsName;
                    if (/[\/\\]render[\/\\]/i.test(fp) || /[\/\\]_grade[\/\\]/i.test(fp)) return L;
                }
            } catch (e) {}
        }
        return null;
    }

    // Return the topmost footage layer that is NOT a render or grade —
    // the "active plate variant" that new imports should cover. Handles
    // the case where the user rendered a processed plate variant
    // (stabilized, denoised, retimed, …) and sits it above the original
    // plate: both are plate-like, and renders/grades should stack above
    // the topmost one so they actually cover the active variant instead
    // of landing between the two plates.
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

    // Find all layers in a comp whose source path contains the given segment
    // (e.g. "render" for {shot}/render/, "_grade" for {shots}/_grade/).
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

    // ── Main ──────────────────────────────────────────────────────────────────
    var shotComps = collectShotComps();
    if (shotComps.length === 0) { alert("No *_comp compositions found in the project."); return; }

    var importedFiles = buildImportedFilesMap();
    var binShots = getBinFolder("Shots");

    var statsImported      = 0;
    var statsSkipped       = 0;
    var statsComps         = 0;
    var compsWithNewRenders = [];
    var dryRunLog          = [];
    var errorLog           = [];

    if (!dryRun) app.beginUndoGroup("Import Renders & Grades");
    try {

        // Two return sources per comp:
        //   render → per-shot folder {shots}/{shot}/render/   (all files)
        //   grade  → flat folder {shots}/_grade/              (files whose name
        //                                                     starts with shot)
        //
        // Processed bottom → top. AE's comp.layers.add() always puts the new
        // layer at the top of the stack, so processing "render" before "grade"
        // guarantees grades end up above VFX returns. Within a category, files
        // are sorted oldest-first so the newest ends up topmost within its group.

        // Pre-read the flat grades folder once per run. Safety-net scaffold
        // in case the user runs this before Shot Roundtrip ever created
        // _grade/. See lib/write_readmes.jsx for the README content.
        var gradesDir = new Folder(fsShots.fsName + "/_grade");
        if (!gradesDir.exists) {
            gradesDir.create();
            try {
                var readmeHelper = new File((new File($.fileName)).parent.parent.fsName + "/lib/write_readmes.jsx");
                if (readmeHelper.exists) $.evalFile(readmeHelper);
                if (typeof writeGradeReadme === "function") writeGradeReadme(gradesDir);
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
            // grade: filter flat _grade/ by filename prefix (shot name).
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
            var plateLayer = findPlateLayer(comp, shotName);
            var shotBin    = getChildBin(binShots, shotName);
            var addedAny   = false;
            var foundAnyCategoryFiles = false;

            for (var catIdx = 0; catIdx < CATEGORIES.length; catIdx++) {
                var category = CATEGORIES[catIdx];
                var catFiles = filesForCategory(category, shotName);
                if (catFiles.length === 0) continue;

                // Oldest first → newest ends up as topmost layer within category.
                catFiles.sort(function (a, b) { return a.modified.getTime() - b.modified.getTime(); });

                foundAnyCategoryFiles = true;

                // Destination bin in the AE project panel. Per-shot sub-bins
                // for both renders and grades so each shot's bin contains
                // everything that belongs to it (plate, render, grade)
                // instead of scattering grades into a separate top-level bin.
                var destBin;
                if (category === "render") {
                    destBin = shotBin ? getOrCreateChildBin(shotBin, "render") : null;
                } else {
                    destBin = shotBin ? getOrCreateChildBin(shotBin, "grade")  : null;
                }

                // Stack anchor: new layers will be moved to just above this
                // layer so imports sit directly above the active plate
                // variant (renders) or above the existing render/grade stack
                // (grades). Anchor updates to each new layer so oldest-first
                // processing ends up with newest on top within the category.
                //   - renders: anchor = topmost plate-like layer (so a
                //     stabilized/denoised plate variant sitting above the
                //     original plate still gets covered, not sandwiched).
                //     Falls back to plateLayer if nothing plate-like is
                //     found (defensive — shouldn't happen in practice).
                //   - grades:  anchor = topmost existing render-or-grade,
                //     or topmost plate-like if none exists yet.
                var plateLikeTop = findTopmostPlateLike(comp) || plateLayer;
                var stackAnchor = null;
                if (category === "render") {
                    stackAnchor = plateLikeTop;
                } else {
                    stackAnchor = findTopmostRenderOrGrade(comp) || plateLikeTop;
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
                            // Cyan (14) for renders to distinguish from the
                            // plate (green 9). Grades stay purple (11).
                            footageItem.label = (category === "grade") ? 11 : 14;
                            importedFiles[fsPath] = footageItem;
                        } catch (e) {
                            errorLog.push(shotName + " [" + category + "]: failed to import " + catFile.name + " — " + e.message);
                            continue;
                        }
                    }

                    try {
                        var newLayer = comp.layers.add(footageItem);
                        // Alignment strategy:
                        //   Anchor to the comp's own "cut in" / "cut out"
                        //   markers (which shot_roundtrip writes onto every
                        //   shotComp) — these are the authoritative cut
                        //   boundaries and are independent of whether the
                        //   plate has been trimmed to cut or is still showing
                        //   full handles. Falls back to plate-layer markers,
                        //   then plate inPoint/outPoint, then comp time 0.
                        //
                        //   Duration-driven branch:
                        //     - new clip duration ≈ cut duration  → cut-only
                        //       render (Resolve default): frame 0 → cut_in.
                        //     - new clip duration ≈ plate duration → full
                        //       file with handles: mirror plate.startTime and
                        //       pin inPoint/outPoint to the cut range.
                        //     - unknown duration → frame 0 → cut_in (safe).
                        var cutInCT = null, cutOutCT = null;
                        try {
                            if (comp.markerProperty && comp.markerProperty.numKeys > 0) {
                                for (var mkI = 1; mkI <= comp.markerProperty.numKeys; mkI++) {
                                    var mkV = comp.markerProperty.keyValue(mkI);
                                    var mkCmt = (mkV && mkV.comment) ? String(mkV.comment).toLowerCase() : "";
                                    if (mkCmt === "cut in")  cutInCT  = comp.markerProperty.keyTime(mkI);
                                    if (mkCmt === "cut out") cutOutCT = comp.markerProperty.keyTime(mkI);
                                }
                            }
                        } catch (eCM) {}
                        if (cutInCT === null && plateLayer) {
                            try {
                                var pLM = plateLayer.property("Marker");
                                if (pLM && pLM.numKeys > 0) {
                                    for (var pmkI = 1; pmkI <= pLM.numKeys; pmkI++) {
                                        var pmkV = pLM.keyValue(pmkI);
                                        var pmkCmt = (pmkV && pmkV.comment) ? String(pmkV.comment).toLowerCase() : "";
                                        // plate-layer markers are in SOURCE time
                                        if (pmkCmt === "cut in")  cutInCT  = plateLayer.startTime + pLM.keyTime(pmkI);
                                        if (pmkCmt === "cut out") cutOutCT = plateLayer.startTime + pLM.keyTime(pmkI);
                                    }
                                }
                            } catch (ePLM) {}
                        }
                        if (cutInCT === null && plateLayer) {
                            cutInCT  = plateLayer.inPoint;
                            cutOutCT = plateLayer.outPoint;
                        }

                        var alignTolerance = 0.5 / comp.frameRate;
                        var newSrcDur = 0;
                        try { newSrcDur = footageItem.duration; } catch (eND) {}

                        if (cutInCT !== null) {
                            var cutDur = (cutOutCT !== null) ? (cutOutCT - cutInCT) : 0;
                            var plateSrcDur = 0;
                            if (plateLayer) { try { plateSrcDur = plateLayer.source.duration; } catch (ePD) {} }

                            if (newSrcDur > 0 && cutDur > 0 &&
                                Math.abs(newSrcDur - cutDur) <= alignTolerance) {
                                newLayer.startTime = cutInCT;
                            } else if (plateLayer && plateSrcDur > 0 &&
                                       Math.abs(newSrcDur - plateSrcDur) <= alignTolerance) {
                                newLayer.startTime = plateLayer.startTime;
                                try {
                                    if (cutOutCT !== null) newLayer.outPoint = cutOutCT;
                                    newLayer.inPoint = cutInCT;
                                } catch (eTrim) {}
                            } else {
                                newLayer.startTime = cutInCT;
                            }
                        } else {
                            newLayer.startTime = 0;
                        }
                        newLayer.position.setValue([comp.width / 2, comp.height / 2]);
                        newLayer.label = (category === "grade") ? 11 : 14;
                        // Copy cut markers onto the new layer. Prefer the plate
                        // layer's own markers; fall back to the comp's
                        // comp-level markers (which shot_roundtrip writes to
                        // every shotComp) so imports still get markers even
                        // when the plate layer was created outside the
                        // roundtrip and has none of its own.
                        var srcMkr = null;
                        if (plateLayer) {
                            try {
                                var pm = plateLayer.property("Marker");
                                if (pm && pm.numKeys > 0) srcMkr = pm;
                            } catch (eMkrPL) {}
                        }
                        if (!srcMkr) {
                            try {
                                if (comp.markerProperty && comp.markerProperty.numKeys > 0) {
                                    srcMkr = comp.markerProperty;
                                }
                            } catch (eMkrCP) {}
                        }
                        if (srcMkr) {
                            try {
                                var newMkr = newLayer.property("Marker");
                                for (var mi = 1; mi <= srcMkr.numKeys; mi++) {
                                    try { newMkr.setValueAtTime(srcMkr.keyTime(mi), srcMkr.keyValue(mi)); } catch (e) {}
                                }
                            } catch (eMkrWrite) {}
                        }
                        // Position the new layer just above the current stack
                        // anchor (plate for renders, topmost existing render/
                        // grade for grades). Then promote the new layer to the
                        // anchor so the next iteration stacks above it — keeps
                        // the "newest on top within category" ordering while
                        // planting the whole group right above the plate.
                        if (stackAnchor) {
                            try { newLayer.moveBefore(stackAnchor); } catch (eMove) {}
                            stackAnchor = newLayer;
                        }
                        addedAny = true;
                        statsImported++;
                    } catch (e) {
                        errorLog.push(shotName + " [" + category + "]: failed to add layer " + catFile.name + " — " + e.message);
                    }
                }
            }

            // If no category folder had any files, nothing to do for this comp.
            if (!foundAnyCategoryFiles) continue;

            statsComps++;
            if (dryRun) continue;

            // Within each category, keep only the topmost (newest) layer enabled.
            var CATEGORY_SEGMENTS = { render: "render", grade: "_grade" };
            for (var ck = 0; ck < CATEGORIES.length; ck++) {
                var catLayers = findLayersByPathSegment(comp, CATEGORY_SEGMENTS[CATEGORIES[ck]]);
                for (var li = 1; li < catLayers.length; li++) {
                    catLayers[li].enabled = false;
                    catLayers[li].audioEnabled = false;
                }
            }

            // Plate enabled-state is intentionally NOT touched. The user
            // manages plate visibility manually so they can verify the
            // import before hiding the source — see the README "Tools →
            // Import Renders & Grades" block for the reasoning.

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
