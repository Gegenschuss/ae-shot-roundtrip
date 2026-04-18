/*
       _____                          __
      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
             /___/

  VFX SHOT ROUNDTRIP
  After Effects ExtendScript
================================================================================

PURPOSE
-------
Extracts selected shots from an After Effects edit and prepares them as
self-contained render plates for VFX work in Nuke (or any other application).
After the VFX render is returned, the script imports it back and wires it into
the live AE comp structure so the result is immediately visible in the edit via
Premiere Pro Dynamic Link.

The source edit in AE stays intact and fully live throughout the entire process.


WHAT HAPPENS PER SHOT
----------------------
For each selected layer the script:

  1.  Creates a  shotname_comp  containing the source footage at its original
      resolution (+ optional overscan), with head and tail handles around the
      cut. The cut always starts at frame 1001 inside the comp.
      A burn-in overlay shows the shot name and current frame number.

  2a. DIRECT FOOTAGE LAYERS:
      All layer attributes (transforms, effects, masks, time remap, opacity,
      blending mode) are moved from the source layer into the _comp layer inside
      a new  shotname_container  comp. The original layer in the main edit is
      replaced by _container, leaving the edit flat and unmodified.

  2b. PRECOMP LAYERS:
      The script traverses into the precomp hierarchy to find the underlying
      footage, respecting time remaps and stretch at every level. It creates
      the _comp there and replaces the footage layer with it. The topmost
      precomp is renamed to  shotname_container. The main edit is untouched.
      If two precomps share the same footage source, only one _comp is created.

  3.  Creates a  shotname_dynamicLink  comp for use in Premiere Pro:
      - Direct footage: wraps _container, duration trimmed to plate length.
      - Precomp: wraps the (renamed) _container precomp at full duration.
      CUT IN / CUT OUT markers indicate the exact cut points.

  4.  Queues  shotname_comp  in the AE render queue. Output: ProRes 4444 .mov
      to  Roundtrip/{shotName}/plate/  (or the configured OM template).

  5.  After render, imports the plate directly from
      Roundtrip/{shotName}/plate/  back into  shotname_comp
      as the VFX return layer (all other layers disabled).

  6.  Writes a Nuke .nk AppendClip master script assembling all plates into a
      timeline, plus one self-contained .nk per shot for individual artist
      handoff (optional).

  7.  Writes a Premiere Pro FCP 7 XML placing all shots on separate tracks,
      aligned to their original edit positions with handle room on both sides
      (optional).


UI OPTIONS
----------
  Prefix          Shot name prefix, e.g. "shot_" → shot_010, shot_020 …
  Start Number    First shot number (padded to 3 digits).
  Increment       Step between shot numbers (default 10).
  Handles         Head and tail handle frames added around the cut.
  Overscan        Expand plate resolution by a percentage (0 = off).
  OM Template     AE Output Module template name for the render (e.g. ProRes 4444).
  Generate Nuke Script      Write the .nk AppendClip script to Roundtrip/.
  Generate Premiere XML     Write the FCP 7 .xml sequence to Roundtrip/.
  Skip Render (debug)       Build the comp structure without triggering a render.
                            When active, render and file-copy are skipped entirely —
                            use this to test comp setup without waiting for a render.

BUTTONS
-------
  Run Roundtrip   Run full prep → render → import → export pipeline.
  Cancel          Dismiss without changes.


HANDLES
-------
Head and tail handles give the VFX artist extra frames for retiming, blending,
or shot extensions. If a clip has less source footage than requested the handles
are clamped and the Roundtrip Complete report lists the affected shots with the
actual frame counts available on each side.


FOLDER STRUCTURE
----------------
The AE project must be saved two levels inside the project root:

    MyProject/
      ae/
        MyProject.aep       ← project file saved here
      Roundtrip/
        {shotName}/
          plate/            ← rendered plate (.mov); re-imported into AE
          render/           ← VFX return goes here (for Nuke artist handoff)
          {shotName}.nk     ← per-shot Nuke script
        _grades/            ← Resolve graded returns (flat, shared)
        dynamicLink/        ← Dynamic Link wrapper comps
        {project}_Comp.nk   ← AppendClip .nk master (optional)
        {project}.xml       ← Premiere FCPXML (optional, overwritten each run)


PREMIERE PRO — DYNAMIC LINK WORKFLOW
--------------------------------------
  1. Import the generated .xml into Premiere (File > Import).
     Each shot appears on its own track, full plate duration, cut aligned to
     its original position. CUT IN / CUT OUT markers show the trim points.

  2. For each shot, right-click the clip > Replace With After Effects Comp.
     Select the matching  shotname_dynamicLink  comp from the AE project.

  3. Trim the relinked clip to the CUT markers. The Dynamic Link comp contains
     the full plate range so trim room is identical to the rendered .mov.

  4. Once the VFX return is imported back by the script, the Dynamic Link
     updates live in Premiere — no re-linking needed.


NOTES
-----
- Run the script with the target comp active and the desired layers selected.
- Direct footage and precomp layers can be mixed in the same selection.
- The AE project must be saved before running (output paths are derived from it).
- If a  shotname_comp  already exists in the project that shot is skipped.
- Bit depth is upgraded to 16-bit if the project is currently set lower.
- Audio-only layers, solids, shapes, and text layers are skipped automatically.
- The Warp Stabilizer warning that AE shows when effects are added via script
  cannot be suppressed — dismiss it with OK. The effect moves correctly.


================================================================================
*/

{
    function vfxRoundtripEpsilon() {
        var proj = app.project;
        if (!proj || !proj.activeItem || !(proj.activeItem instanceof CompItem)) { alert("Shot Roundtrip: please open a composition first."); return; }
        if (proj.activeItem.selectedLayers.length === 0) { alert("Shot Roundtrip: select at least one layer."); return; }
        // Required downstream (output paths, render save, per-shot .nk filenames).
        // Guard here so we fail before collecting shot data, not mid-process.
        if (!proj.file) { alert("Shot Roundtrip: save the project first (output paths are derived from the .aep location)."); return; }

        // ------------------------------------------------
        // UI
        // ------------------------------------------------
        var LABEL_W = 100; var FIELD_H = 22;

        var addRow = function(parent, labelText) {
            var g = parent.add("group");
            g.orientation = "row"; g.alignChildren = ["left", "center"]; g.spacing = 8;
            var lbl = g.add("statictext", undefined, labelText);
            lbl.preferredSize = [LABEL_W, FIELD_H];
            return g;
        };

        var dlg = new Window("dialog", "Gegenschuss VFX Roundtrip");
        dlg.orientation = "column"; dlg.alignChildren = ["fill", "top"];
        dlg.spacing = 10; dlg.margins = 14;

        // ── Shot Settings ──────────────────────────────
        var pnlMain = dlg.add("panel", undefined, "Shot Settings");
        pnlMain.orientation = "column"; pnlMain.alignChildren = ["fill", "top"];
        pnlMain.spacing = 6; pnlMain.margins = [10, 15, 10, 10];

        var r1 = addRow(pnlMain, "Prefix:");
        var etPrefix = r1.add("edittext", undefined, "shot_");
        etPrefix.preferredSize = [120, FIELD_H];

        var r2 = addRow(pnlMain, "Start Number:");
        var etStartNum = r2.add("edittext", undefined, "010");
        etStartNum.preferredSize = [60, FIELD_H];

        var r2b = addRow(pnlMain, "Increment:");
        var etIncrement = r2b.add("edittext", undefined, "10");
        etIncrement.preferredSize = [60, FIELD_H];

        var r3 = addRow(pnlMain, "Handles:");
        var etHandles = r3.add("edittext", undefined, "50");
        etHandles.preferredSize = [60, FIELD_H];
        r3.add("statictext", undefined, "frames");

        // ── Pipeline Options ───────────────────────────
        var pnlOpt = dlg.add("panel", undefined, "Pipeline Options");
        pnlOpt.orientation = "column"; pnlOpt.alignChildren = ["fill", "top"];
        pnlOpt.spacing = 6; pnlOpt.margins = [10, 15, 10, 10];

        var chkCreateNuke     = pnlOpt.add("checkbox", undefined, "Create Nuke Scripts");  chkCreateNuke.value     = true;
        var chkExportXML      = pnlOpt.add("checkbox", undefined, "Export Shot XML");    chkExportXML.value      = true;
        var chkCreateDynLink  = pnlOpt.add("checkbox", undefined, "Create dynamicLink Comps"); chkCreateDynLink.value = true;

        var r4 = addRow(pnlOpt, "Overscan:");
        var etOverscan = r4.add("edittext", undefined, "10");
        etOverscan.preferredSize = [50, FIELD_H];
        r4.add("statictext", undefined, "%");

        var r5 = addRow(pnlOpt, "OM Template:");
        var etOM = r5.add("edittext", undefined, "ProRes 4444");
        etOM.preferredSize = [150, FIELD_H];

        var r6 = addRow(pnlOpt, "Shots Folder:");
        var etShotsFolder = r6.add("edittext", undefined, "../Roundtrip");
        etShotsFolder.preferredSize = [150, FIELD_H];

        var chkSkipRender     = pnlOpt.add("checkbox", undefined, "Skip Render (debug)");           chkSkipRender.value     = false;

        // ── Buttons ────────────────────────────────────
        var btnGrp = dlg.add("group");
        btnGrp.orientation = "row"; btnGrp.alignment = ["fill", "bottom"]; btnGrp.margins = [0, 4, 0, 0];
        var btnCancel  = btnGrp.add("button", undefined, "Cancel");        btnCancel.preferredSize  = [80,  28];
        var btnSpacer  = btnGrp.add("statictext", undefined, "");          btnSpacer.alignment      = ["fill", "center"];
        var btnOk      = btnGrp.add("button", undefined, "Run Roundtrip"); btnOk.preferredSize      = [130, 28];

        btnOk.onClick     = function() { dlg.close(1); };
        btnCancel.onClick = function() { dlg.close(2); };

        var dlgResult = dlg.show();
        if (dlgResult === 2) return;

        // ------------------------------------------------
        // HELPER UTILITIES
        // ------------------------------------------------
        function waitForFile(fObj, maxRetries) {
            if (fObj.exists) return true;
            var attempts = 0;
            while (attempts < maxRetries) { $.sleep(500); fObj = new File(fObj.fsName); if (fObj.exists) return true; attempts++; }
            return false;
        }
        function reportError(phase, e, hint) {
            var msg = "Error in " + phase + " Phase:\n" + e.message + "\nLine: " + e.line;
            if (hint) msg += "\n\n" + hint;
            alert(msg);
        }

        // Non-modal progress palette so the user can tell the script is still
        // alive during the long prep + post-render phases. Returns an object
        // with update/close. If palette creation fails for any reason, all
        // methods become no-ops — the roundtrip runs the same, just silent.
        // Palette is invisible during AE's renderQueue.render() since that
        // call blocks the script thread; AE's own render window takes over
        // the UI during that phase.
        function makeProgressPanel() {
            var win = null, statusTxt = null, detailTxt = null, bar = null, alive = false;
            var cancelled = false, btnCancel = null;
            try {
                win = new Window("palette", "Gegenschuss \u00b7 Shot Roundtrip", undefined, { resizeable: false });
                win.orientation = "column";
                win.alignChildren = ["fill", "top"];
                win.margins = 14; win.spacing = 6;
                win.preferredSize.width = 440;
                statusTxt = win.add("statictext", undefined, "Starting roundtrip\u2026");
                statusTxt.alignment = ["fill", "top"];
                detailTxt = win.add("statictext", undefined, " ");
                detailTxt.alignment = ["fill", "top"];
                bar = win.add("progressbar", undefined, 0, 100);
                bar.preferredSize = [410, 8];
                var btnGrp = win.add("group");
                btnGrp.orientation = "row"; btnGrp.alignment = ["right", "top"];
                btnCancel = btnGrp.add("button", undefined, "Cancel");
                btnCancel.preferredSize = [90, 24];
                btnCancel.onClick = function () {
                    cancelled = true;
                    try { btnCancel.text = "Cancelling\u2026"; btnCancel.enabled = false; win.update(); } catch (eCl) {}
                };
                win.show();
                alive = true;
            } catch (eProg) { alive = false; }
            return {
                update: function (status, detail, pct) {
                    if (!alive) return;
                    try {
                        if (typeof status === "string") statusTxt.text = status;
                        if (typeof detail === "string") detailTxt.text = (detail.length ? detail : " ");
                        if (typeof pct === "number") bar.value = Math.max(0, Math.min(100, pct));
                        win.update();
                    } catch (e) {}
                },
                isCancelled: function () { return cancelled; },
                close: function () {
                    if (!alive) return;
                    alive = false;
                    try { win.close(); } catch (e) {}
                }
            };
        }
        // No-op default so any early-return paths before the real palette is
        // created don't crash on progress.update/close/isCancelled calls.
        var progress = { update: function(){}, isCancelled: function(){ return false; }, close: function(){} };

        // Cancel-check helper. Call at the top of each long loop iteration:
        //   if (cancelCheck()) return;
        // The return triggers any enclosing try/finally, which correctly ends
        // the active undo group, so the user can Cmd/Ctrl+Z to roll back
        // partial work. Doesn't work during renderQueue.render() — that call
        // blocks the script and AE's own render window takes over, including
        // its own cancel button.
        function cancelCheck() {
            if (!progress.isCancelled()) return false;
            progress.close();
            alert("Roundtrip cancelled by user.\n\nIf any comps or renders were already created, use Cmd/Ctrl+Z to undo them.");
            return true;
        }
        function pad(n, s) { var str = "" + n; while (str.length < s) str = "0" + str; return str; }
        function getBinFolder(name) { for (var i=1;i<=proj.numItems;i++) { if(proj.item(i).name==name && proj.item(i) instanceof FolderItem) return proj.item(i); } return proj.items.addFolder(name); }
        // "cut in"/"cut out" markers created here use protectedRegion so a
        // later time remap or stretch on the layer keeps the marker locked
        // to its frame rather than drifting in time.
        function cutMarker(comment) { var m = new MarkerValue(comment); m.protectedRegion = true; return m; }

        // Reversed clips deserve a loud, deliberate dialog — they bake
        // reversed frames into the plate, which breaks camera tracking,
        // motion vectors, particle trails, smoke, debris, and any other
        // direction-sensitive VFX work downstream.
        //
        // Non-reversed time effects (forward ramps, non-100% stretches) do
        // NOT pop a dialog. The Confirm Shots list mentions them inline so
        // the user can eyeball, but they're ignorable — auto-precompose
        // handles top-level ones, and nested non-reversed effects just mean
        // "the artist gets a plate with the ramp baked in," which is usually
        // the intent.
        //
        // Returns true if the user wants to proceed, false if they cancel.
        function confirmReversedClips(reversed) {
            var w = new Window("dialog", "\u26A0  Reversed clips detected — manual check required");
            w.orientation = "column"; w.alignChildren = ["fill", "top"];
            w.spacing = 10; w.margins = 14;

            w.add("statictext", undefined,
                reversed.length + " reversed clip" + (reversed.length === 1 ? "" : "s") +
                " " + (reversed.length === 1 ? "is" : "are") +
                " in your selection (top-level or nested):");

            var lb = w.add("listbox", undefined, [], {
                numberOfColumns: 3, showHeaders: true,
                columnTitles: ["layer", "path", "effect"],
                columnWidths: [240, 520, 220]
            });
            for (var k = 0; k < reversed.length; k++) {
                var row = lb.add("item", reversed[k].layerName);
                row.subItems[0].text = reversed[k].path;
                row.subItems[1].text = reversed[k].label;
            }
            lb.preferredSize = [1000, Math.min(reversed.length * 22 + 40, 420)];

            var warn = w.add("statictext", undefined,
                "Reversed plates break camera tracking, motion vectors, and direction-sensitive VFX \u2014 particle trails, smoke, fire, debris, splashes all look obviously wrong when played backwards. Please verify each one really should ship reversed.",
                { multiline: true });
            warn.preferredSize = [1000, 60];

            var btnGrp = w.add("group");
            btnGrp.alignment = ["right", "bottom"];
            btnGrp.margins = [0, 4, 0, 0];
            var btnCancel   = btnGrp.add("button", undefined, "Cancel \u2014 I'll fix first");
            var btnContinue = btnGrp.add("button", undefined, "Continue \u2014 reversed is intentional");
            btnCancel.preferredSize   = [180, 28];
            btnContinue.preferredSize = [240, 28];
            btnCancel.onClick   = function () { w.close(2); };
            btnContinue.onClick = function () { w.close(1); };

            return w.show() === 1;
        }

        // Precompose Trimmed — auto-precompose layers with time remap or stretch.
        // Same algorithm as helpers/precompose_trimmed.jsx:
        // descending index order, move all attributes, full comp duration, trim to original in/out.
        // Returns true on success, false on failure.
        function autoPrecomposeTrimmed(comp, layerInfos) {
            // Snapshot non-affected selection by layer name+index pair. precompose()
            // collapses comp.selectedLayers to just the most-recently-created
            // precomp layer, so anything the user had selected that wasn't in
            // layerInfos would silently drop out — visible as a missing selection
            // in the later Confirm Shots list. After the loop we re-select those
            // originals plus every new precomp layer we just created.
            var trIndices = {};
            for (var ii = 0; ii < layerInfos.length; ii++) trIndices[layerInfos[ii].index] = true;
            var preservedSelection = [];
            for (var pi = 1; pi <= comp.numLayers; pi++) {
                var pl = comp.layer(pi);
                if (pl.selected && !trIndices[pi]) preservedSelection.push(pl);
            }
            var newPrecompLayers = [];

            layerInfos.sort(function(a, b) { return b.index - a.index; });
            app.beginUndoGroup("Auto-Precompose for Roundtrip");
            var aAutoPCBin = null; // lazily created on first precompose
            try {
                for (var ap = 0; ap < layerInfos.length; ap++) {
                    if (progress.isCancelled()) { app.endUndoGroup(); return false; }
                    var aInfo = layerInfos[ap];
                    try { progress.update(null,
                        "layer " + (ap + 1) + " of " + layerInfos.length + ": " + aInfo.name,
                        9 + 2 * (ap / Math.max(1, layerInfos.length))); } catch(eProgAP) {}

                    var aBaseName = aInfo.name + "_precomp";
                    var aName = aBaseName;
                    var aSuffix = 1;
                    while (true) {
                        var aExists = false;
                        for (var ak = 1; ak <= proj.numItems; ak++) { if (proj.item(ak).name === aName) { aExists = true; break; } }
                        if (!aExists) break;
                        aName = aBaseName + "_" + aSuffix; aSuffix++;
                    }

                    var aNewComp = comp.layers.precompose([aInfo.index], aName, true);
                    aNewComp.duration = comp.duration;

                    // File the new precomp into /Shots/autoPrecomps/ instead of
                    // leaving it stranded at the project root alongside whatever
                    // the user had organised there.
                    try {
                        if (!aAutoPCBin) aAutoPCBin = getShotBin(getBinFolder("Shots"), "autoPrecomps");
                        aNewComp.parentFolder = aAutoPCBin;
                    } catch (eBin) {}

                    var aPrecompLayer = null;
                    for (var aj = 1; aj <= comp.numLayers; aj++) {
                        if (comp.layer(aj).source === aNewComp) { aPrecompLayer = comp.layer(aj); break; }
                    }
                    if (aPrecompLayer) {
                        aPrecompLayer.inPoint  = aInfo.inPoint;
                        aPrecompLayer.outPoint = aInfo.outPoint;
                        newPrecompLayers.push(aPrecompLayer);
                    }
                }
            } catch(eAP) {
                alert("Auto-precompose failed:\n" + eAP.message + "\nLine: " + eAP.line);
                app.endUndoGroup();
                return false;
            }

            // Restore the combined selection: preserved originals + every new precomp.
            try {
                for (var dsel = 1; dsel <= comp.numLayers; dsel++) comp.layer(dsel).selected = false;
                for (var psi = 0; psi < preservedSelection.length; psi++) {
                    try { preservedSelection[psi].selected = true; } catch(ePs) {}
                }
                for (var npi = 0; npi < newPrecompLayers.length; npi++) {
                    try { newPrecompLayers[npi].selected = true; } catch(eNp) {}
                }
            } catch(eSelPC) {}

            app.endUndoGroup();
            return true;
        }

        function getShotBin(parentBin, shotName) {
            for (var i=1;i<=proj.numItems;i++) { if(proj.item(i).name===shotName && proj.item(i) instanceof FolderItem && proj.item(i).parentFolder===parentBin) return proj.item(i); }
            var f = proj.items.addFolder(shotName); f.parentFolder = parentBin; f.label = 2; return f; // Yellow — shot bin
        }

        // Returns { start, end } in source footage time for a direct footage layer,
        // accounting for time remap or stretch so handles are added at the right place.
        function getRequiredSourceRange(layer) {
             try {
                var startT = 0;
                var endT = 0;
                if (layer.timeRemapEnabled) {
                    var tr = layer.property("Time Remap");
                    if (tr.numKeys >= 1) {
                        // Walk every key; min/max of their values is the true source range the
                        // remap covers. Avoids valueAtTime's linear extrapolation past the last
                        // key, which inflates range.end for reversed-stretch conversions (last
                        // key sits at compEnd - frameDur, not at layer.outPoint).
                        startT = tr.keyValue(1);
                        endT   = tr.keyValue(1);
                        for (var ki = 2; ki <= tr.numKeys; ki++) {
                            var kv = tr.keyValue(ki);
                            if (kv < startT) startT = kv;
                            if (kv > endT)   endT   = kv;
                        }
                    } else {
                        startT = tr.valueAtTime(layer.inPoint,  false);
                        endT   = tr.valueAtTime(layer.outPoint, false);
                    }
                } else {
                    var factor = 100 / layer.stretch;
                    startT = (layer.inPoint - layer.startTime) * factor;
                    endT = (layer.outPoint - layer.startTime) * factor;
                }
                if (startT > endT) { var temp = startT; startT = endT; endT = temp; }
                return { start: startT, end: endT };
            } catch(e) { return { start: 0, end: 0 }; }
        }

        // Maps a time in the containing comp to a time in the layer's source,
        // respecting time remap when enabled and stretch otherwise. Relies on the
        // keys having LINEAR interpolation so valueAtTime extrapolates past the
        // first/last key (which is what the handle-extended layer actually plays);
        // convertStretchReversalToRemap forces that interpolation type explicitly.
        function mapTimeToSource(layer, compTime) {
            if (layer.timeRemapEnabled) {
                try { return layer.property("ADBE Time Remapping").valueAtTime(compTime, false); } catch(e) {}
            }
            var stretch = (layer.stretch !== 0) ? layer.stretch : 100;
            return (compTime - layer.startTime) * (100 / stretch);
        }

        // Recursively searches comp for ALL footage layers anywhere in the hierarchy.
        // No time filtering — finds every footage file regardless of which frames are
        // currently visible. Returns an array of { footageLayer, footageComp }.
        // The source time range for each layer is calculated separately via
        // getRequiredSourceRange(), which reads the layer's own in/out and time remap.
        function findAllFootageInPrecomp(comp, path) {
            var currentPath = (path || []).concat([comp.name]);
            var results = [];
            for (var li = 1; li <= comp.numLayers; li++) {
                var l = comp.layer(li);
                if (!l.hasVideo || l.guideLayer || l.adjustmentLayer || l.nullLayer) continue;
                if (l.source === null) continue;
                // Footage file?
                var isFile = false;
                try { if (l.source.mainSource && l.source.mainSource.file) isFile = true; } catch(e) {}
                if (isFile) {
                    results.push({ footageLayer: l, footageComp: comp, breadcrumb: currentPath });
                } else if (l.source instanceof CompItem) {
                    // Sub-precomp: recurse and collect all footage inside
                    var sub = findAllFootageInPrecomp(l.source, currentPath);
                    for (var si = 0; si < sub.length; si++) results.push(sub[si]);
                }
            }
            return results;
        }

        // --- NUKE SCRIPT WRITER ---
        // ── Nuke Write node delivery settings ─────────────────────────────────────
        // Edit these if your pipeline uses a different codec or color transform.
        var NUKE_FILE_TYPE = "mov64";
        var NUKE_CODEC     = "appr";                  // Apple ProRes
        var NUKE_PROFILE   = "ProRes 4:4:4:4 12-bit";
        var NUKE_COLOR     = "rec709";

        // Nuke script_directory expression. When set as Root.project_directory,
        // Nuke resolves every relative Read/Write path against the .nk file's
        // own folder at load time — so the whole Roundtrip tree can be moved,
        // zipped, or handed off without breaking any paths.
        var NUKE_PROJECT_DIR = "\"\\[python \\{nuke.script_directory()\\}\\]\"";

        // Return absPath relative to baseDir if absPath is inside baseDir;
        // otherwise return absPath unchanged (absolute fallback).
        function relativizePath(absPath, baseDir) {
            var a = String(absPath).replace(/\\/g, "/");
            var b = String(baseDir).replace(/\\/g, "/");
            if (b.charAt(b.length - 1) !== "/") b += "/";
            if (a.indexOf(b) === 0) return a.substring(b.length);
            return a;
        }

        function writeNukeScript(nukeFile, nukeData, globalFPS) {
            nukeData.sort(function(a, b) { return a.globalStartTime - b.globalStartTime; });

            var globalMinFrame = 9999999;  // sentinel — clamped by first shot's plate start
            var globalMaxFrame = -9999999; // sentinel — clamped by last shot's plate end
            for(var k=0; k<nukeData.length; k++) {
                var d = nukeData[k];
                var fs = 1001 - d.handles; 
                var fe = fs + d.fullDurationFrames - 1; 
                if (fs < globalMinFrame) globalMinFrame = fs;
                if (fe > globalMaxFrame) globalMaxFrame = fe;
            }
            if (globalMinFrame > globalMaxFrame) { globalMinFrame = 1001; globalMaxFrame = 1100; }

            var projW = 2048; var projH = 1556; // fallback: DCI 2K — overridden by actual plate size below
            if (nukeData.length > 0) { projW = nukeData[0].w; projH = nukeData[0].h; }

            var nukeBaseDir = nukeFile.parent.fsName;

            var nk = "";
            nk += "# Generated by Gegenschuss \u00b7 AE Shot Roundtrip\n";
            nk += "# https://github.com/Gegenschuss/ae-shot-roundtrip\n\n";
            nk += "Root {\n inputs 0\n name " + nukeFile.fsName + "\n";
            nk += " project_directory " + NUKE_PROJECT_DIR + "\n";
            nk += " format \"" + projW + " " + projH + " 0 0 " + projW + " " + projH + " 1 ProjectFormat\"\n";
            nk += " fps " + globalFPS + "\n";
            nk += " first_frame " + globalMinFrame + "\n";
            nk += " last_frame " + globalMaxFrame + "\n";
            nk += " lock_range true\n";
            nk += " colorManagement Nuke\n";
            nk += " OCIO_config nuke-default\n";
            nk += "}\n";

            var spacingX = 1000; var stepY = 400; var inputCounter = 0;

            for (var i = nukeData.length - 1; i >= 0; i--) {
                var d = nukeData[i];
                var platePath  = relativizePath(d.platePath,  nukeBaseDir);
                var renderPath = relativizePath(d.renderPath, nukeBaseDir);
                var shot = d.name;
                var x = i * spacingX;
                var y = 0;

                if (i < nukeData.length - 1) {
                    var nextShot = nukeData[i+1];
                    var currentEndFrame = Math.round((d.globalStartTime * globalFPS) + d.cutDurationFrames);
                    var nextStartFrame = Math.round(nextShot.globalStartTime * globalFPS);
                    var gapFrames = nextStartFrame - currentEndFrame;

                    if (gapFrames > 0) {
                        nk += "Constant {\n inputs 0\n channels rgba\n color {0 0 0 0}\n name Gap_After_" + shot + "\n xpos " + (x + (spacingX/2)) + "\n ypos " + (y + stepY) + "\n}\n";
                        nk += "FrameRange {\n first_frame 1\n last_frame " + gapFrames + "\n name GapDuration_" + i + "\n xpos " + (x + (spacingX/2)) + "\n ypos " + (y + (stepY*2)) + "\n}\n";
                        inputCounter++;
                    }
                }

                var fullStart = 1001 - d.handles;
                var fullEnd = fullStart + d.fullDurationFrames - 1;
                var cutStart = 1001;
                var cutEnd = 1001 + d.cutDurationFrames - 1;
                var offsetVal = fullStart - 1;

                nk += "Read {\n inputs 0\n file \"" + platePath + "\"\n format \"" + d.w + " " + d.h + " 0 0 " + d.w + " " + d.h + " 1\"\n first 1\n last " + d.fullDurationFrames + "\n origfirst 1\n origlast " + d.fullDurationFrames + "\n name Read_" + shot + "\n xpos " + x + "\n ypos " + y + "\n}\n";
                nk += "TimeOffset {\n time_offset " + offsetVal + "\n name Offset_" + shot + "\n xpos " + x + "\n ypos " + (y + stepY) + "\n}\n";
                nk += "FrameRange {\n first_frame " + cutStart + "\n last_frame " + cutEnd + "\n name Range_" + shot + "\n xpos " + x + "\n ypos " + (y + (stepY*2)) + "\n}\n";
                nk += "Write {\n render_order " + (i+1) + "\n file \"" + renderPath + "\"\n file_type " + NUKE_FILE_TYPE + "\n mov64_codec " + NUKE_CODEC + "\n mov_prores_codec_profile \"" + NUKE_PROFILE + "\"\n colorspace " + NUKE_COLOR + "\n channels rgba\n use_limit 1\n first " + fullStart + "\n last " + fullEnd + "\n name Write_" + shot + "\n xpos " + x + "\n ypos " + (y + (stepY*3)) + "\n}\n";
                inputCounter++; 
            }

            var masterX = ((nukeData.length * spacingX) / 2);
            var masterY = stepY * 4.5;
            nk += "AppendClip {\n inputs " + inputCounter + "\n name Append_Timeline\n xpos " + masterX + "\n ypos " + masterY + "\n}\n";
            nk += "Viewer {\n inputs 0\n viewerProcess \"rec709\"\n name Viewer1\n xpos " + masterX + "\n ypos " + (masterY + 600) + "\n}\n";

            if (!nukeFile.open("w")) {
                alert("Failed to write master Nuke script:\n" + nukeFile.fsName + "\n\nCheck folder permissions and free disk space.");
                return;
            }
            nukeFile.write(nk);
            nukeFile.close();
        }


        // --- PER-SHOT NUKE SCRIPT WRITER ---
        function writeNukeShotScript(shotFile, d, globalFPS) {
            var fullStart = 1001 - d.handles;
            var fullEnd   = fullStart + d.fullDurationFrames - 1;
            var cutStart  = 1001;
            var cutEnd    = 1001 + d.cutDurationFrames - 1;
            var offsetVal = fullStart - 1;

            var shotBaseDir = shotFile.parent.fsName;
            var platePath  = relativizePath(d.platePath,  shotBaseDir);
            var renderPath = relativizePath(d.renderPath, shotBaseDir);

            var nk = "";
            nk += "# ============================================================\n";
            nk += "# Gegenschuss \u00b7 Per-shot Nuke handoff (AE Shot Roundtrip)\n";
            nk += "# https://github.com/Gegenschuss/ae-shot-roundtrip\n";
            nk += "# ============================================================\n";
            nk += "#\n";
            nk += "# Shot:      " + d.name + "\n";
            nk += "# Format:    " + d.w + "x" + d.h + " @ " + globalFPS + " fps\n";
            nk += "# Handles:   " + d.handles + "f head + " + d.handles + "f tail\n";
            nk += "# Frames:    full " + fullStart + "-" + fullEnd + "  (" + d.fullDurationFrames + "f, with handles)\n";
            nk += "#            cut  " + cutStart + "-" + cutEnd + "  (" + d.cutDurationFrames + "f, editorial)\n";
            nk += "#\n";
            nk += "# Graph:  Read (plate, full range) -> TimeOffset -> FrameRange (cut)\n";
            nk += "#         -> Write -> " + renderPath + "\n";
            nk += "#\n";
            nk += "# HANDOFF RULES\n";
            nk += "# -------------\n";
            nk += "# 1. Render filename MUST start with \"" + d.name + "\" followed by a\n";
            nk += "#    non-alphanumeric character (e.g. \"" + d.name + "_v01.mov\",\n";
            nk += "#    \"" + d.name + "_comp_v03.mov\"). The AE \"Import Renders & Grades\"\n";
            nk += "#    tool matches returns to shot comps by this filename prefix.\n";
            nk += "# 2. Do NOT rename this shot's folder or the render/ subfolder.\n";
            nk += "#    The AE pipeline expects {shots}/" + d.name + "/render/.\n";
            nk += "# 3. Render the full handle range (" + fullStart + "-" + fullEnd + "), not just\n";
            nk += "#    the editorial cut. Extra frames on each side are used for\n";
            nk += "#    retime, blend, and late edit changes back in Premiere.\n";
            nk += "# ============================================================\n\n";
            nk += "Root {\n inputs 0\n name " + shotFile.fsName + "\n";
            nk += " project_directory " + NUKE_PROJECT_DIR + "\n";
            nk += " format \"" + d.w + " " + d.h + " 0 0 " + d.w + " " + d.h + " 1 ProjectFormat\"\n";
            nk += " fps " + globalFPS + "\n";
            nk += " first_frame " + fullStart + "\n";
            nk += " last_frame " + fullEnd + "\n";
            nk += " lock_range true\n colorManagement Nuke\n OCIO_config nuke-default\n}\n";

            nk += "Read {\n inputs 0\n file \"" + platePath + "\"\n";
            nk += " format \"" + d.w + " " + d.h + " 0 0 " + d.w + " " + d.h + " 1\"\n";
            nk += " first 1\n last " + d.fullDurationFrames + "\n";
            nk += " origfirst 1\n origlast " + d.fullDurationFrames + "\n";
            nk += " name Read_" + d.name + "\n xpos 0\n ypos 0\n}\n";

            nk += "TimeOffset {\n time_offset " + offsetVal + "\n name Offset_" + d.name + "\n xpos 0\n ypos 400\n}\n";
            nk += "FrameRange {\n first_frame " + cutStart + "\n last_frame " + cutEnd + "\n name Range_" + d.name + "\n xpos 0\n ypos 800\n}\n";

            nk += "Write {\n render_order 1\n file \"" + renderPath + "\"\n";
            nk += " file_type " + NUKE_FILE_TYPE + "\n mov64_codec " + NUKE_CODEC + "\n";
            nk += " mov_prores_codec_profile \"" + NUKE_PROFILE + "\"\n";
            nk += " colorspace " + NUKE_COLOR + "\n";
            nk += " channels rgba\n use_limit 1\n first " + fullStart + "\n last " + fullEnd + "\n";
            nk += " name Write_" + d.name + "\n xpos 0\n ypos 1200\n}\n";

            nk += "Viewer {\n inputs 0\n viewerProcess \"rec709\"\n name Viewer1\n xpos 200\n ypos 1200\n}\n";

            if (!shotFile.open("w")) {
                alert("Failed to write per-shot Nuke script:\n" + shotFile.fsName + "\n\nCheck folder permissions and free disk space.");
                return;
            }
            shotFile.write(nk);
            shotFile.close();
        }



        function addGuideBurnIn(comp, shotName, cutFrame1001, fullRenderStart, handleFrames) {
            var txtLayer = comp.layers.byName("GUIDE_BURNIN");
            if (!txtLayer) { txtLayer = comp.layers.addText(shotName); txtLayer.name = "GUIDE_BURNIN"; }
            txtLayer.locked = false; txtLayer.guideLayer = true; txtLayer.label = 16; txtLayer.moveToBeginning();
            var txtProp = txtLayer.property("Source Text");
            var txtDoc = txtProp.value; txtDoc.fontSize = 80; txtDoc.fillColor = [0.5, 0.5, 0.5];
            txtDoc.justification = ParagraphJustification.LEFT_JUSTIFY; txtProp.setValue(txtDoc);
            txtLayer.property("Anchor Point").setValue([0,0]); txtLayer.position.setValue([100, 200]);
            var expr = "sName = '" + shotName + "';\r" + "cutFrame = " + cutFrame1001 + ";\r" + "fullStart = " + fullRenderStart + ";\r" + "handles = " + handleFrames + ";\r" + "\r" + "renderStartFrame = cutFrame - handles;\r" + "timeSinceStart = time - fullStart;\r" + "framesSinceStart = Math.round(timeSinceStart / thisComp.frameDuration);\r" + "finalFrame = renderStartFrame + framesSinceStart;\r" + "\r" + "sName + '\\r' + 'Frame: ' + finalFrame;";
            txtLayer.property("Source Text").expression = expr; txtLayer.locked = true;
        }

        // Mirror the comp's "cut in"/"cut out" markers onto the locked
        // GUIDE_BURNIN layer. AE prevents editing marker keys on a locked
        // layer, so this gives the cut markers UI-level protection on top of
        // the protectedRegion flag (which only guards against time-remap drift).
        function addGuideBurnInMarkers(comp, cutInSec, cutOutSec) {
            var gl = comp.layers.byName("GUIDE_BURNIN");
            if (!gl) return;
            gl.locked = false;
            var mp = gl.property("Marker");
            mp.setValueAtTime(cutInSec,  cutMarker("cut in"));
            mp.setValueAtTime(cutOutSec, cutMarker("cut out"));
            gl.locked = true;
        }

        // ------------------------------------------------
        // EXECUTION — collect layers, build comps, render
        // ------------------------------------------------
        if (!proj || !proj.activeItem) return;
        var mainComp = proj.activeItem;
        if (proj.bitsPerChannel < 16) {
            if (!confirm("Project is currently " + proj.bitsPerChannel + "-bit. Upgrade to 16-bit for this render?\n(This permanently changes the project.)")) return;
            proj.bitsPerChannel = 16;
        }

        var shotPrefix = etPrefix.text;
        var startNum = parseInt(etStartNum.text, 10);
        var handleFrames = parseInt(etHandles.text, 10);
        var overscanPercent = parseFloat(etOverscan.text); if(isNaN(overscanPercent)) overscanPercent=0;
        var omTemplate = etOM.text;
        var useNukeStart = true;
        var increment = parseInt(etIncrement.text, 10);
        if (isNaN(increment) || increment < 1) increment = 10;

        var selLayers = [];
        var skippedLayers = [];

        var rawSelection = mainComp.selectedLayers;
        for (var i = 0; i < rawSelection.length; i++) {
            var l = rawSelection[i];
            if (!l.hasVideo || l.guideLayer || l.adjustmentLayer || l.nullLayer) continue;
            if (l.source === null) { skippedLayers.push(l.name + " (Shape/Text)"); continue; }
            var isFile = false;
            try { if (l.source.mainSource && l.source.mainSource.file) isFile = true; } catch(e) {}
            var isPC = (l.source instanceof CompItem);
            if (isFile || isPC) {
                selLayers.push({ layer: l, isPrecomp: isPC, mainLayerIdx: l.index });
            } else {
                skippedLayers.push(l.name + " (Solid/Shape/Text)");
            }
        }

        selLayers.sort(function(a, b) { return a.layer.index - b.layer.index; });

        progress = makeProgressPanel();
        progress.update("Preflight: scanning " + selLayers.length + " selected layer(s)\u2026", "", 2);

        // ── Preflight: time remap or time stretch ─────────────────────────────
        // Scan not just the selected (top-level) layers but the full precomp
        // hierarchy below them. A reversed precomp nested two levels deep
        // (container_comp → reversed_comp → footage) silently bakes into the
        // plate render — so the user has to see it before we proceed.
        function isTimeRemapReversed(layer) {
            if (!layer.timeRemapEnabled) return false;
            try {
                var tr   = layer.property("Time Remap");
                var vIn  = tr.valueAtTime(layer.inPoint,  false);
                var vOut = tr.valueAtTime(layer.outPoint, false);
                return vOut < vIn;
            } catch (eTR) { return false; }
        }
        function describeTimeEffect(layer) {
            if (layer.timeRemapEnabled) {
                var trRev = isTimeRemapReversed(layer);
                return { label: "[time remap" + (trRev ? ", reversed" : "") + "]", reversed: trRev };
            }
            if (Math.abs(layer.stretch - 100) > 0.01) {
                var stRev = layer.stretch < 0;
                return { label: "[time stretch " + Math.round(layer.stretch) + "%" + (stRev ? ", reversed" : "") + "]", reversed: stRev };
            }
            return null;
        }
        // Only layers whose time effect could realistically bake reversed or
        // ramped pixels into the plate are scan-relevant. Skips disabled,
        // guide, null, adjustment, and non-AV (text/shape) layers — a
        // "reversed" text layer isn't a tracking problem.
        function isScanRelevantLayer(l) {
            if (!l.enabled || l.guideLayer || l.nullLayer || l.adjustmentLayer) return false;
            if (!(l instanceof AVLayer)) return false;
            if (!l.hasVideo) return false;
            return true;
        }
        function walkPrecompForEffects(comp, path, results) {
            for (var li = 1; li <= comp.numLayers; li++) {
                var l = comp.layer(li);
                if (!isScanRelevantLayer(l)) continue;
                var fx = describeTimeEffect(l);
                if (fx) results.push({ layerName: l.name, path: path.join(" \u203A "), label: fx.label, reversed: fx.reversed, topLevel: false });
                if (l.source instanceof CompItem) {
                    walkPrecompForEffects(l.source, path.concat([l.source.name]), results);
                }
            }
        }

        // Scan each selected layer + every precomp beneath it. Reversed
        // entries drive the loud dialog; non-reversed entries just become
        // passive info in the Confirm Shots listing later.
        function scanSelLayerForEffects(tl) {
            var results = [];
            var topFx = describeTimeEffect(tl);
            if (topFx) results.push({ layerName: tl.name, path: mainComp.name, label: topFx.label, reversed: topFx.reversed, topLevel: true });
            if (tl.source instanceof CompItem) {
                walkPrecompForEffects(tl.source, [mainComp.name, tl.source.name], results);
            }
            return results;
        }

        // Convert negative-stretch reversals to equivalent reversed time remaps.
        // A reversed time remap survives the roundtrip cleanly (mapTimeToSource
        // reads valueAtTime directly, so source-time resolution is data-driven
        // rather than formula-dependent). Negative stretch does not: multiple
        // "Hold in Place" options all map source time differently, and a single
        // general formula has been unreliable in practice. So we rewrite the
        // reversal into the kind the pipeline handles, before scanning.
        //
        // AE's default Hold in Place = Layer In-point, which this assumes. If
        // the user had a different Hold in Place option, the visual extent may
        // shift slightly after `stretch = 100` — ExtendScript doesn't expose
        // the anchor mode, so this is the best we can do without prompting.
        function convertStretchReversalToRemap(layer) {
            try {
                if (layer.timeRemapEnabled) return false;
                if (!(layer.stretch < 0)) return false;
                var srcDur = (layer.source && layer.source.duration) ? layer.source.duration : 0;
                if (srcDur <= 0) return false;

                // For negatively-stretched layers AE swaps the reported in/out:
                //   inPoint  = comp-LATER  edge (source-earlier side)
                //   outPoint = comp-EARLIER edge (source-later side)
                // Capture, compute source times, normalize into comp-timeline order.
                var startT   = layer.startTime;
                var stretch  = layer.stretch;
                var rawIn    = layer.inPoint;
                var rawOut   = layer.outPoint;
                var frameDur = layer.containingComp.frameDuration;

                // sourceTime(compTime) = (compTime - startTime) * (100 / stretch).
                // For reversed layers AE anchors startTime one frame past the last
                // rendered source frame, so subtract one frameDur to land on the
                // source time that's actually on screen.
                var srcAtRawIn  = (rawIn  - startT) * (100 / stretch) - frameDur;
                var srcAtRawOut = (rawOut - startT) * (100 / stretch) - frameDur;

                var compStart, compEnd, srcAtStart, srcAtEnd;
                if (rawIn <= rawOut) {
                    compStart  = rawIn;        compEnd  = rawOut;
                    srcAtStart = srcAtRawIn;   srcAtEnd = srcAtRawOut;
                } else {
                    compStart  = rawOut;       compEnd  = rawIn;
                    srcAtStart = srcAtRawOut;  srcAtEnd = srcAtRawIn;
                }

                if (srcAtStart < 0)      srcAtStart = 0;
                if (srcAtStart > srcDur) srcAtStart = srcDur;
                if (srcAtEnd   < 0)      srcAtEnd   = 0;
                if (srcAtEnd   > srcDur) srcAtEnd   = srcDur;

                // Reset stretch to forward; anchor the layer back to the original
                // comp range (AE may have relocated it).
                layer.stretch   = 100;
                layer.startTime = compStart;
                layer.inPoint   = compStart;
                layer.outPoint  = compEnd;

                // Enable time remap. Select layer + property to dodge the "hidden
                // property" error that setValueAtTime throws otherwise. Access the
                // property via matchname.
                layer.selected = true;
                layer.timeRemapEnabled = true;
                var tr = layer.property("ADBE Time Remapping");
                tr.selected = true;

                // Two keys placed at the handle boundaries (not the cut
                // boundaries) so AE never has to extrapolate past the last key
                // in the *_dynamicLink wrapper or anywhere else the layer gets
                // extended beyond the cut. Earlier attempts placed keys at the
                // cut boundaries and relied on LINEAR extrapolation; that broke
                // once moveLayerAttribs copied the keys into containerInner at
                // a comp-time that didn't line up with the layer's visible
                // range, and the handles froze on a single source frame.
                //
                // Slope comes from the stretch-derived source times at compStart
                // and endKeyTime, so it captures the actual speed regardless of
                // the original negative-stretch magnitude (-100 → slope -1,
                // -50 → -2, -200 → -0.5, -124 → ≈ -0.806, etc.). Values clamp
                // to [0, srcDur] — when the source runs out, clamping freezes
                // on the source edge, the physically correct behaviour.
                //
                // Trade-off of keeping only 2 keys: when clamping kicks in
                // (cut sits within handleSec of a source edge, in the direction
                // the handle extends into the source), the single linear segment
                // between the clamped boundary key and the unclamped one
                // slightly warps the cut-range playback too. In that rare edge
                // case, the cut plays at a slightly shifted speed. Adding back
                // interior keys at compStart/endKeyTime would pin the cut range
                // exactly; left out here for timeline cleanliness.
                var endKeyTime = compEnd - frameDur;
                var keySpan    = endKeyTime - compStart;
                var handleSec  = (layer.containingComp && layer.containingComp.frameRate)
                               ? (handleFrames / layer.containingComp.frameRate)
                               : 0;
                var slope      = (keySpan > 0) ? ((srcAtEnd - srcAtStart) / keySpan) : 0;
                var preTime    = compStart  - handleSec;
                var postTime   = endKeyTime + handleSec;
                var preVal     = srcAtStart - slope * handleSec;
                var postVal    = srcAtEnd   + slope * handleSec;
                if (preVal  < 0)      preVal  = 0;
                if (preVal  > srcDur) preVal  = srcDur;
                if (postVal < 0)      postVal = 0;
                if (postVal > srcDur) postVal = srcDur;

                // Write our keys first, then prune AE's auto-keys — clearing
                // them before writing can leave the property in an unusable state.
                tr.setValueAtTime(preTime,  preVal);
                tr.setValueAtTime(postTime, postVal);

                for (var k = tr.numKeys; k >= 1; k--) {
                    var kt = tr.keyTime(k);
                    if (Math.abs(kt - preTime)  > 0.0005 &&
                        Math.abs(kt - postTime) > 0.0005) {
                        tr.removeKey(k);
                    }
                }

                // Force LINEAR interpolation on both keys. AE's default for a
                // fresh setValueAtTime on time remap is BEZIER, which eases
                // between keys — wrong for constant-speed reversed playback.
                for (var ki = 1; ki <= tr.numKeys; ki++) {
                    try {
                        tr.setInterpolationTypeAtKey(ki, KeyframeInterpolationType.LINEAR, KeyframeInterpolationType.LINEAR);
                    } catch (eInterp) {}
                }

                return true;
            } catch (e) { return false; }
        }
        function convertAllStretchReversalsInComp(comp) {
            var n = 0;
            for (var li = 1; li <= comp.numLayers; li++) {
                var l = comp.layer(li);
                if (!isScanRelevantLayer(l)) continue;
                if (convertStretchReversalToRemap(l)) n++;
                if (l.source instanceof CompItem) {
                    n += convertAllStretchReversalsInComp(l.source);
                }
            }
            return n;
        }
        // Snapshot the mainComp selection — our converter flips layer/property
        // selection flags to dodge the "hidden property" error, which collapses
        // the user's original selection in the UI. Restore it after the loop.
        var preConvertSelection = [];
        for (var pcs = 0; pcs < mainComp.selectedLayers.length; pcs++) {
            preConvertSelection.push(mainComp.selectedLayers[pcs]);
        }
        var autoConverted = 0;
        progress.update("Converting any reversed clips to time remaps\u2026", "0 of " + selLayers.length, 5);
        for (var cri = 0; cri < selLayers.length; cri++) {
            if (cancelCheck()) return;
            progress.update(null,
                "layer " + (cri + 1) + " of " + selLayers.length,
                5 + 2 * (cri / Math.max(1, selLayers.length)));
            var crl = selLayers[cri].layer;
            if (!isScanRelevantLayer(crl)) continue;
            if (convertStretchReversalToRemap(crl)) autoConverted++;
            if (crl.source instanceof CompItem) {
                autoConverted += convertAllStretchReversalsInComp(crl.source);
            }
        }
        try {
            for (var dse = 1; dse <= mainComp.numLayers; dse++) mainComp.layer(dse).selected = false;
            for (var rse = 0; rse < preConvertSelection.length; rse++) {
                try { preConvertSelection[rse].selected = true; } catch(eSelRestore) {}
            }
        } catch(eSelSnap) {}

        var trAffected   = []; // top-level, fed to autoPrecomposeTrimmed
        var reversedList = []; // drives the loud confirm dialog
        for (var ti = 0; ti < selLayers.length; ti++) {
            var tl = selLayers[ti].layer;
            var topFx = describeTimeEffect(tl);
            if (topFx) {
                trAffected.push({ selIdx: ti, index: tl.index, name: tl.name, inPoint: tl.inPoint, outPoint: tl.outPoint,
                    label: topFx.label, reversed: topFx.reversed });
            }
            var layerEffects = scanSelLayerForEffects(tl);
            for (var le = 0; le < layerEffects.length; le++) {
                if (layerEffects[le].reversed) reversedList.push(layerEffects[le]);
            }
        }

        // Loud dialog only when a reversal is present. Non-reversed time
        // effects proceed silently and show up as passive info next to the
        // shot in the Confirm Shots dialog.
        if (reversedList.length > 0) {
            // Close palette before modal so macOS hands focus to the dialog.
            progress.close();
            var revOk = confirmReversedClips(reversedList);
            progress = makeProgressPanel();
            if (!revOk) { progress.close(); return; } // user canceled
            progress.update("Reversed clips confirmed, continuing\u2026", "", 8);
        }
        if (trAffected.length > 0) {
            progress.update("Auto-precomposing " + trAffected.length + " time-remapped layer(s)\u2026", "", 9);
            if (!autoPrecomposeTrimmed(mainComp, trAffected)) {
                // autoPrecomposeTrimmed returns false on internal error OR on
                // cancel. Surface the cancel message if that's why we stopped.
                if (cancelCheck()) return;
                progress.close(); return;
            }

            // Re-scan the selection — the precomposed layers replaced the originals.
            // Apply the same filter as the initial scan: auto-precompose's selection
            // restoration puts every pre-existing selection back on the layer panel,
            // including shape/text/null/adjustment/guide layers the user happened to
            // have selected alongside the real shots. Without this filter those come
            // through as roundtrip candidates and crash downstream on null .source.
            selLayers = [];
            for (var ri = 1; ri <= mainComp.numLayers; ri++) {
                if (!mainComp.layer(ri).selected) continue;
                var rLayer = mainComp.layer(ri);
                if (!rLayer.hasVideo || rLayer.guideLayer || rLayer.adjustmentLayer || rLayer.nullLayer) continue;
                if (rLayer.source === null) { skippedLayers.push(rLayer.name + " (Shape/Text)"); continue; }
                var rIsFile = false;
                try { if (rLayer.source.mainSource && rLayer.source.mainSource.file) rIsFile = true; } catch(eRIS) {}
                var rIsPrecomp = (rLayer.source instanceof CompItem);
                if (rIsFile || rIsPrecomp) {
                    selLayers.push({ layer: rLayer, isPrecomp: rIsPrecomp, mainLayerIdx: rLayer.index });
                } else {
                    skippedLayers.push(rLayer.name + " (Solid/Shape/Text)");
                }
            }
            if (selLayers.length === 0) {
                alert("No layers selected after auto-precompose. Please select the precomposed layers and try again.");
                progress.close();
                return;
            }
        }

        // Expand precomp entries: each footage layer found inside becomes its own shot.
        // Direct footage entries pass through unchanged.
        progress.update("Expanding precomps and resolving source ranges\u2026", "0 of " + selLayers.length, 11);
        var expandedLayers = [];
        for (var ei = 0; ei < selLayers.length; ei++) {
            if (cancelCheck()) return;
            progress.update(null, (ei + 1) + " of " + selLayers.length + " selected layers walked", 11 + (5 * ei / Math.max(1, selLayers.length)));
            var eItem = selLayers[ei];
            if (!eItem.isPrecomp) {
                expandedLayers.push({
                    layer: eItem.layer, mainLayerIdx: eItem.layer.index, isPrecomp: false, found: null, totalInPrecomp: 0,
                    // Snapshot the source id NOW, before any replaceSource runs later in the
                    // processing loop — dedup must key on the original source, not whatever
                    // shotComp has taken its place.
                    originalSourceId: (eItem.layer.source && eItem.layer.source.id) ? eItem.layer.source.id : null
                });
            } else {
                var eFounds = findAllFootageInPrecomp(eItem.layer.source);
                if (eFounds.length === 0) {
                    skippedLayers.push(eItem.layer.name + " (no footage found inside precomp)");
                } else {
                    for (var ef = 0; ef < eFounds.length; ef++) {
                        var eF = eFounds[ef];
                        // Derive the source range from the precomp layer's cut bounds in mainComp
                        // (not the footage layer's own in/out, which is often the full untrimmed duration).
                        // For footage directly inside the precomp we do a two-step mapping:
                        //   mainComp cut → precomp time (via the precomp layer's stretch/remap)
                        //   precomp time → source time   (via the footage layer's stretch/remap)
                        // For deeper nesting we fall back to the footage layer's own in/out.
                        var eRange;
                        if (eF.footageComp === eItem.layer.source) {
                            var pcIn  = mapTimeToSource(eItem.layer, eItem.layer.inPoint);
                            var pcOut = mapTimeToSource(eItem.layer, eItem.layer.outPoint);
                            var rA = mapTimeToSource(eF.footageLayer, pcIn);
                            var rB = mapTimeToSource(eF.footageLayer, pcOut);
                            eRange = { start: rA, end: rB };
                        } else {
                            eRange = getRequiredSourceRange(eF.footageLayer);
                        }
                        // Normalize so start <= end. A reversed time remap (decreasing
                        // values from in to out) or a nested reversal produces an inverted
                        // range here; downstream math — plate duration, handle expansion,
                        // marker placement — all assume forward order. Reversal direction
                        // is tracked separately via scanLayerTimeEffect/reversedList, so
                        // losing the ordering here is safe.
                        if (eRange.start > eRange.end) {
                            var eRtmp = eRange.start; eRange.start = eRange.end; eRange.end = eRtmp;
                        }
                        expandedLayers.push({
                            layer: eItem.layer, mainLayerIdx: eItem.layer.index, isPrecomp: true, totalInPrecomp: eFounds.length,
                            found: { footageLayer: eF.footageLayer, footageComp: eF.footageComp,
                                     rangeStart: eRange.start, rangeEnd: eRange.end,
                                     breadcrumb: eF.breadcrumb || [] },
                            // Snapshot the original footage source id NOW. The processing loop
                            // later calls replaceSource on this footageLayer — after that,
                            // reading footageLayer.source.id returns the shotComp id, so dedup
                            // must key on this pre-mutation snapshot instead.
                            originalSourceId: (eF.footageLayer.source && eF.footageLayer.source.id) ? eF.footageLayer.source.id : null
                        });
                    }
                }
            }
        }

        if (expandedLayers.length === 0) {
            var msg = "Keine gültigen Layer gefunden.";
            if (skippedLayers.length > 0) msg += "\n" + skippedLayers.length + " Layer wurden übersprungen.";
            alert(msg); progress.close(); return;
        }

        // ── Confirm found shots before doing anything ──────────────────────────
        // Pre-scan existing comp names so we can flag shots that will be skipped
        var confExisting = {};
        for (var ei = 1; ei <= proj.numItems; ei++) {
            if (proj.item(ei) instanceof CompItem) confExisting[proj.item(ei).name] = true;
        }

        var cfFps = mainComp.frameRate;

        // Default per-shot overscan flag from the global UI field
        for (var cfi = 0; cfi < expandedLayers.length; cfi++) {
            expandedLayers[cfi].overscan = false;
        }

        // Each row: { cols:[shot, frames, res, notice, source, os], layerIdx }
        // layerIdx=-1 for skip rows (shot comp already exists).
        var confRows = [];
        var confTotalFrames = 0;
        var confNameMaxLen   = 0;
        var confFramesMaxLen = 0;
        var confResMaxLen    = 0;
        var confNoticeMaxLen = 0;
        var confPathMaxLen   = 0;

        progress.update("Preparing Confirm Shots dialog\u2026", "0 of " + expandedLayers.length, 12);
        try {
        for (var cfi = 0; cfi < expandedLayers.length; cfi++) {
            if (cancelCheck()) return;
            // Per-row update so a hang/crash surfaces the exact offending row.
            progress.update(null, "row " + (cfi + 1) + " of " + expandedLayers.length, 12 + (2 * cfi / Math.max(1, expandedLayers.length)));
            var cfNum  = startNum + cfi * increment;
            var cfName = shotPrefix + pad(cfNum, 3);
            var cfItem = expandedLayers[cfi];

            // Already exists?
            if (!!confExisting[cfName + "_comp"]) {
                confRows.push({ cols: [cfName, "", "", "\u2014 skip", "already exists", ""], layerIdx: -1 });
                if (cfName.length > confNameMaxLen) confNameMaxLen = cfName.length;
                continue;
            }

            // Cut duration
            var cfCutFrames = Math.round((cfItem.layer.outPoint - cfItem.layer.inPoint) * cfFps);
            confTotalFrames += cfCutFrames + handleFrames * 2;

            // FPS mismatch?
            var cfSrcFps = 0;
            try {
                cfSrcFps = cfItem.isPrecomp
                    ? cfItem.found.footageLayer.source.frameRate
                    : cfItem.layer.source.frameRate;
            } catch(eFps) {}
            var cfFpsMismatch = (cfSrcFps > 0 && Math.abs(cfSrcFps - cfFps) > 0.01);

            // Source resolution (no overscan here — overscan shown via toggle column)
            var cfSrcW = 0, cfSrcH = 0;
            try {
                var cfSrc = cfItem.isPrecomp ? cfItem.found.footageLayer.source : cfItem.layer.source;
                cfSrcW = cfSrc.width; cfSrcH = cfSrc.height;
            } catch(eRes) {}

            // Breadcrumb path — shorten long chains to "first > … > last"
            var cfPath;
            if (cfItem.isPrecomp) {
                var cfSrcName = (cfItem.found.footageLayer.source && cfItem.found.footageLayer.source.name) || "(no source)";
                var cfParts = cfItem.found.breadcrumb.concat([cfSrcName]);
                if (cfParts.length > 3) {
                    cfPath = cfParts[0] + " > \u2026 > " + cfParts[cfParts.length - 1];
                } else {
                    cfPath = cfParts.join(" > ");
                }
            } else {
                cfPath = (cfItem.layer.source && cfItem.layer.source.name) || "(no source)";
            }

            var cfFrames = cfCutFrames + "f";
            var cfRes    = (cfSrcW > 0) ? (cfSrcW + "\u00d7" + cfSrcH) : "";

            // Notice column: fps mismatch + time-effect tags. Reversed effects are
            // already surfaced via the loud dialog at the top of the run; we repeat
            // them here per-shot so it's clear which shots the warning covered.
            var cfNoticeParts = [];
            if (cfFpsMismatch) cfNoticeParts.push("!!! fps " + cfSrcFps + "\u2260" + cfFps);
            var cfEffects = scanSelLayerForEffects(cfItem.layer);
            for (var cfe = 0; cfe < cfEffects.length; cfe++) {
                cfNoticeParts.push(cfEffects[cfe].label);
            }
            var cfNotice = cfNoticeParts.join("  ");

            var cfOsMark = cfItem.overscan ? "\u2715" : "";
            confRows.push({ cols: [cfName, cfFrames, cfRes, cfNotice, cfPath, cfOsMark], layerIdx: cfi });
            if (cfName.length   > confNameMaxLen)   confNameMaxLen   = cfName.length;
            if (cfFrames.length > confFramesMaxLen) confFramesMaxLen = cfFrames.length;
            if (cfRes.length    > confResMaxLen)    confResMaxLen    = cfRes.length;
            if (cfNotice.length > confNoticeMaxLen) confNoticeMaxLen = cfNotice.length;
            if (cfPath.length   > confPathMaxLen)   confPathMaxLen   = cfPath.length;
        }
        } catch (eConfRows) {
            try { progress.close(); } catch(eCl){}
            alert("Preparing Confirm Shots dialog failed at row " + (typeof cfi === "number" ? (cfi + 1) : "?") + ":\n\n" +
                  eConfRows.message + "\nLine: " + eConfRows.line);
            return;
        }

        var screenW = 1920;
        try { screenW = $.screens[0].right - $.screens[0].left; } catch(e) {}
        var maxDlgW = Math.round(screenW * 0.9);

        var confColShot   = Math.max(confNameMaxLen   * 9 + 24, 80);
        var confColFrames = Math.max(confFramesMaxLen * 8 + 16, 60);
        var confColRes    = Math.max(confResMaxLen    * 8 + 16, 90);
        var confColNotice = Math.max(confNoticeMaxLen * 7 + 16, 80);
        var confColOs     = 32;
        var confColSource = Math.max(confPathMaxLen   * 8 + 24, 260);
        // Source column gets whatever space remains after the fixed columns.
        var confFixedW = confColShot + confColFrames + confColRes + confColNotice + confColOs + 60;
        confColSource  = Math.min(confColSource, maxDlgW - confFixedW);
        confColSource  = Math.max(confColSource, 260);
        var confDlgW   = Math.min(confFixedW + confColSource, maxDlgW);
        confDlgW = Math.max(confDlgW, 800);

        var confDlg = new Window("dialog", "Confirm Shots");
        confDlg.orientation = "column"; confDlg.alignChildren = ["fill", "top"];
        confDlg.spacing = 8; confDlg.margins = 14;

        // Header info
        var confProjName = proj.file ? proj.file.displayName.replace(".aep", "") : "unsaved project";
        var confInfoTxt = confProjName + "   \u2022   " + mainComp.name + "   \u2022   " + cfFps + " fps   \u2022   handles: " + handleFrames + "f";
        if (overscanPercent > 0) confInfoTxt += "   \u2022   +" + overscanPercent + "% overscan (toggle per shot with \u2715)";
        confDlg.add("statictext", undefined, confInfoTxt);

        if (chkSkipRender.value) {
            confDlg.add("statictext", undefined, "!!! SKIP RENDER IS ON \u2014 comps will be built but nothing will be rendered or imported");
        }

        confDlg.add("statictext", undefined, expandedLayers.length + " shot" + (expandedLayers.length !== 1 ? "s" : "") + " will be created:");

        var confLB = confDlg.add("listbox", undefined, [], {
            multiselect: true,
            numberOfColumns: 6,
            showHeaders: true,
            columnTitles: ["shot", "frames", "res", "notice", "source", "os"],
            columnWidths: [confColShot, confColFrames, confColRes, confColNotice, confColSource, confColOs]
        });
        var confLBH = Math.max(confRows.length * 22 + 40, 200);
        confLBH = Math.min(confLBH, 600);
        confLB.preferredSize = [confDlgW, confLBH];
        for (var cfi = 0; cfi < confRows.length; cfi++) {
            var cfRow = confLB.add("item", confRows[cfi].cols[0]);
            cfRow.subItems[0].text = confRows[cfi].cols[1]; // frames
            cfRow.subItems[1].text = confRows[cfi].cols[2]; // res
            cfRow.subItems[2].text = confRows[cfi].cols[3]; // notice
            cfRow.subItems[3].text = confRows[cfi].cols[4]; // source
            cfRow.subItems[4].text = confRows[cfi].cols[5]; // os
        }

        // Footer + toggle button
        var confFooterTxt = "Total: " + confTotalFrames + "f  (" + (Math.round(confTotalFrames / cfFps * 10) / 10) + "s)   handles: " + handleFrames + "f";
        confDlg.add("statictext", undefined, confFooterTxt);

        var confBtnGrp = confDlg.add("group");
        confBtnGrp.orientation = "row"; confBtnGrp.alignment = ["fill", "bottom"];
        var confSpacer  = confBtnGrp.add("statictext", undefined, ""); confSpacer.alignment = ["fill", "center"];
        var confToggleOs = confBtnGrp.add("button", undefined, "Toggle Overscan"); confToggleOs.preferredSize = [130, 28];
        var confCancel  = confBtnGrp.add("button", undefined, "Cancel");  confCancel.preferredSize  = [80,  28];
        var confOk      = confBtnGrp.add("button", undefined, "Process"); confOk.preferredSize      = [100, 28];

        function toggleOverscanSelection() {
            var sel = confLB.selection;
            if (!sel) return;
            var selArr = (sel.length !== undefined) ? sel : [sel];
            // Collect indices before clearing selection
            var selIndices = [];
            for (var si = 0; si < selArr.length; si++) selIndices.push(selArr[si].index);
            // Toggle data
            for (var si = 0; si < selIndices.length; si++) {
                var ri = confRows[selIndices[si]];
                if (ri.layerIdx < 0) continue;
                expandedLayers[ri.layerIdx].overscan = !expandedLayers[ri.layerIdx].overscan;
                confLB.items[selIndices[si]].subItems[4].text = expandedLayers[ri.layerIdx].overscan ? "\u2715" : "";
            }
            // Force ScriptUI to repaint the listbox — it won't update subItem text
            // while items are selected, so briefly clear and restore selection.
            confLB.selection = null;
            for (var si = 0; si < selIndices.length; si++) confLB.items[selIndices[si]].selected = true;
        }
        confToggleOs.onClick = toggleOverscanSelection;
        // X key on the dialog (more reliable than on the listbox in AE's ScriptUI)
        try {
            confDlg.addEventListener("keydown", function(e) {
                if (e.keyName === "X") { e.preventDefault(); toggleOverscanSelection(); }
            });
        } catch(eKD) {}
        confCancel.onClick = function() { confDlg.close(2); };
        confOk.onClick     = function() { confDlg.close(1); };

        // Palette windows can keep AE's main window from handing focus to a
        // new modal on macOS — the dialog opens behind and looks like a
        // freeze. Close the progress palette before .show() and recreate it
        // after the dialog closes so the Confirm Shots dialog always comes
        // to front.
        progress.update("Waiting for Confirm Shots dialog\u2026", "If nothing visible, check behind AE's main window.", 14);
        progress.close();
        var confResult = confDlg.show();
        progress = makeProgressPanel();
        progress.update("Confirm Shots dialog closed, preparing build\u2026", "", 15);
        if (confResult !== 1) { progress.close(); return; }

        var aepFolder = proj.file.parent;
        var fsShots = new Folder(aepFolder.fsName + "/" + etShotsFolder.text);
        if (!fsShots.exists) fsShots.create();
        if (!fsShots.exists) { alert("Could not create shots folder:\n" + fsShots.fsName); progress.close(); return; }
        var fsScripts = fsShots;

        // Scaffold the Roundtrip/ and _grades/ README.txt files so the
        // handoff tree is self-documenting from day one. See
        // lib/write_readmes.jsx for the content. Both writes are non-fatal
        // (README is a nicety, not a requirement).
        try {
            var readmeHelper = new File((new File($.fileName)).parent.parent.fsName + "/lib/write_readmes.jsx");
            if (readmeHelper.exists) $.evalFile(readmeHelper);
            if (typeof writeRoundtripReadme === "function") writeRoundtripReadme(fsShots);
            var fsGrades = new Folder(fsShots.fsName + "/_grades");
            if (!fsGrades.exists) fsGrades.create();
            if (typeof writeGradesReadme === "function") writeGradesReadme(fsGrades);
        } catch (eRM) { /* non-fatal */ }

        var binShots      = getBinFolder("Shots");

        var renderItems = [];
        var nukeDataList = [];
        var clampedShots = [];
        var stats = { count: 0, origWidth: 0, plateWidth: 0, fps: mainComp.frameRate, bpc: proj.bitsPerChannel };

        app.beginUndoGroup("Roundtrip Prep");
        try {
            // Tracks footage sources already processed (source.id → { bin }) so that
            // if two precomps reference the same footage file, only one _comp is made.
            var processedSourceIds = {};
            var precompLayerRegistry    = {}; // mainLayerIdx → firstName (for multi-footage rename)
            var precompRangeBinRegistry = {}; // mainLayerIdx → FolderItem (per-range bin under /Shots)

            // Pre-build set of existing comp names so the already-processed check is O(1)
            // rather than scanning all project items for every shot.
            var existingCompNames = {};
            for (var ei = 1; ei <= proj.numItems; ei++) {
                if (proj.item(ei) instanceof CompItem) existingCompNames[proj.item(ei).name] = true;
            }

            for (var i = 0; i < expandedLayers.length; i++) {
                if (cancelCheck()) return;
                var item      = expandedLayers[i];
                var layer     = item.layer;
                var isPrecomp = item.isPrecomp;
                var currentNum = startNum + (i * increment);
                var shotName   = shotPrefix + pad(currentNum, 3);
                progress.update(
                    "Building shot comps\u2026",
                    "shot " + (i + 1) + " of " + expandedLayers.length + ": " + shotName,
                    15 + 45 * (i / Math.max(1, expandedLayers.length))
                );
                var osSuffix   = (item.overscan && overscanPercent > 0) ? "_OS" : "";

                // ── Resolve source, range, and footage location ────────────────
                var source, range, footageLayer, footageComp;
                if (isPrecomp) {
                    // item.found was pre-computed during layer expansion
                    source       = item.found.footageLayer.source;
                    range        = { start: item.found.rangeStart, end: item.found.rangeEnd };
                    footageLayer = item.found.footageLayer;
                    footageComp  = item.found.footageComp;
                    // Same footage source already processed by a prior entry?
                    // Check against the ORIGINAL source id captured at expansion time —
                    // a prior iteration's replaceSource may have swapped this footage
                    // layer's live source to a shotComp, which would defeat this check
                    // if we read source.id directly here.
                    if (item.originalSourceId && processedSourceIds[item.originalSourceId]) {
                        continue;
                    }
                } else {
                    source = layer.source;
                    range  = getRequiredSourceRange(layer);
                    // No dedup for direct-footage selections: if the user picked the layer,
                    // they want their own shot for it — even if the same source happens to
                    // live inside a previously-processed precomp. Dedup is strictly a
                    // precomp→precomp concern (see the isPrecomp branch above).
                }

                // Skip if shot comp already exists in project
                if (existingCompNames[shotName + "_comp"] || existingCompNames[shotName + "_comp_OS"]) { skippedLayers.push(shotName + " (already processed — skipped)"); continue; }

                var shotBin = getShotBin(binShots, shotName);
                if (i === 0) stats.origWidth = source.width;

                // ── Overscan sizing ────────────────────────────────────────────
                var shotOverscan = (item.overscan && overscanPercent > 0) ? overscanPercent : 0;
                var osWidth = source.width; var osHeight = source.height;
                if (shotOverscan > 0) {
                    var f = 1 + (shotOverscan / 100);
                    osWidth  = Math.ceil(source.width  * f);
                    osHeight = Math.ceil(source.height * f);
                    if (osWidth  % 2 !== 0) osWidth++;
                    if (osHeight % 2 !== 0) osHeight++;
                }
                if (i === 0) { stats.plateWidth = osWidth; stats.plateHeight = osHeight; }

                var safeDuration = source.duration;
                if (safeDuration <= 0) safeDuration = mainComp.duration;
                var safeFPS = source.frameRate;
                if (safeFPS <= 0) safeFPS = mainComp.frameRate;

                // ── 1. Shot Comp ───────────────────────────────────────────────
                // One extra frame added so the epsilon on timeSpanStart never reaches the comp boundary
                var shotComp = proj.items.addComp(shotName + "_comp" + osSuffix, osWidth, osHeight, source.pixelAspect, safeDuration + (1 / safeFPS), safeFPS);
                shotComp.displayStartTime = 0;
                shotComp.parentFolder = shotBin;
                shotComp.label = 11; // Orange — shot comp (render target)

                var plateInner = shotComp.layers.add(source);
                plateInner.startTime = 0;
                plateInner.position.setValue([osWidth / 2, osHeight / 2]);

                var handleSec = handleFrames / shotComp.frameRate;

                // Clamp to source duration (not shotComp.duration, which has the buffer frame).
                // Both bounds are guarded — downstream shotComp.workAreaStart/Duration must stay
                // inside the comp, and upstream mapTimeToSource can return values outside [0, srcDur]
                // for nested time effects (reversed remaps, layers with startTime offset, etc.).
                var rawStart = range.start - handleSec;
                if (rawStart < 0) rawStart = 0;
                if (rawStart > safeDuration) rawStart = safeDuration;
                var rawEnd = range.end + handleSec;
                if (rawEnd < 0) rawEnd = 0;
                if (rawEnd > safeDuration) rawEnd = safeDuration;
                if ((rawEnd - rawStart) < 0.04) rawEnd = Math.min(rawStart + 0.04, safeDuration);
                if (rawEnd < rawStart) { var _swp = rawStart; rawStart = rawEnd; rawEnd = _swp; }

                var snappedStartFrame  = Math.round(rawStart * safeFPS);
                var fullStart          = snappedStartFrame / safeFPS;
                var snappedEndFrame    = Math.round(rawEnd  * safeFPS);
                var fullDurationFrames = snappedEndFrame - snappedStartFrame;
                var fullDurationSec    = fullDurationFrames / safeFPS;

                var cutStartFrame         = Math.round(range.start * safeFPS);
                var actualLeadingHandles  = cutStartFrame - snappedStartFrame;
                var actualTrailingHandles = snappedEndFrame - Math.round(range.end * safeFPS);
                if (actualLeadingHandles < handleFrames || actualTrailingHandles < handleFrames) {
                    clampedShots.push(shotName + "  (head: " + actualLeadingHandles + "f  tail: " + actualTrailingHandles + "f  /  requested: " + handleFrames + "f, black padding added)");
                }

                var cutStart    = fullStart + handleSec;
                if (fullStart === 0 && (range.start - handleSec < 0)) cutStart = range.start;
                var cutDuration = range.end - range.start;
                if (cutStart + cutDuration > safeDuration) cutDuration = safeDuration - cutStart;

                shotComp.workAreaStart    = fullStart;
                shotComp.workAreaDuration = fullDurationSec;
                shotComp.markerProperty.setValueAtTime(cutStart, cutMarker("cut in"));
                shotComp.markerProperty.setValueAtTime(cutStart + cutDuration, cutMarker("cut out"));

                var cutFrame = useNukeStart ? 1001 : 0;
                if (true) addGuideBurnIn(shotComp, shotName, cutFrame, fullStart, actualLeadingHandles);
                addGuideBurnInMarkers(shotComp, cutStart, cutStart + cutDuration);

                // ── Path-specific: precomp vs. direct footage ──────────────────
                var containerDurFrames;
                if (isPrecomp) {

                    // If the topmost precomp layer has a time remap or speed stretch,
                    // auto-precompose it into a neutral wrapper so _dynamicLink is frame-accurate.
                    // startTime is intentionally excluded: a non-zero startTime is normal layer
                    // positioning and is handled correctly by mapTimeToSource without wrapping.
                    var needsWrap = layer.timeRemapEnabled ||
                                    Math.abs(layer.stretch - 100) > 0.01;
                    if (needsWrap) {
                        // Capture original timing before precompose invalidates the ref —
                        // AE's native precompose does not reliably preserve inPoint/outPoint
                        // on the new mainComp layer, so we restore them ourselves.
                        var origInPointPC  = layer.inPoint;
                        var origOutPointPC = layer.outPoint;
                        var layerIdx = layer.index; // read before precompose invalidates the ref
                        try { mainComp.layers.precompose([layerIdx], shotName + "_container" + osSuffix, true); } catch(ePC) {}
                        // Always re-fetch outside the try — whether precompose succeeded or failed,
                        // mainComp.layer(idx) is valid: it's the wrapper on success, the original on failure.
                        layer = mainComp.layer(layerIdx);
                        // Restore original in/out so the edit's cut placement survives.
                        try {
                            layer.startTime = 0;
                            layer.inPoint   = origInPointPC;
                            layer.outPoint  = origOutPointPC;
                        } catch (eTimingPC) {}
                        // Single-shot precomp → shot bin. Multi-footage container's
                        // per-range bin is handled by the rename block below.
                        if (item.totalInPrecomp === 1) {
                            try { layer.source.parentFolder = shotBin; } catch(ePF) {}
                        }
                        // Patch remaining entries that share the same original layer index.
                        // We compare mainLayerIdx (an integer saved before any processing) instead of
                        // comparing AE layer object references — ExtendScript throws "Object is invalid"
                        // when the invalidated DOM object is evaluated even in a === comparison.
                        for (var upd = i + 1; upd < expandedLayers.length; upd++) {
                            if (expandedLayers[upd].mainLayerIdx === layerIdx) {
                                expandedLayers[upd].layer = layer;
                            }
                        }
                        // Same patch for selLayers so the dynamicLink build loop
                        // later doesn't iterate stale refs ("Object is invalid").
                        for (var sUpdP = 0; sUpdP < selLayers.length; sUpdP++) {
                            if (selLayers[sUpdP].mainLayerIdx === layerIdx) {
                                selLayers[sUpdP].layer = layer;
                            }
                        }
                    } else {
                        // No timing offsets — rename and move the precomp for consistency
                        // (single-footage only — multi-footage is handled by the range-bin
                        //  block below so we don't overwrite the name on intermediate passes)
                        if (item.totalInPrecomp === 1) {
                            try { layer.source.name = shotName + "_container" + osSuffix; } catch(eRn) {}
                            try { layer.source.parentFolder = shotBin; } catch(ePF2) {}
                        }
                    }

                    // Multi-footage precomps get their own range bin "/Shots/shot_FIRST_LAST/"
                    // and the container source lives inside it. Runs on every matching entry
                    // so the range (and bin name) grows as we encounter later footage — the
                    // last write wins, which after the loop gives "shot_FIRST_LAST_container"
                    // in "/Shots/shot_FIRST_LAST/".
                    if (item.totalInPrecomp > 1) {
                        var lIdx = item.mainLayerIdx;
                        if (!precompLayerRegistry[lIdx]) {
                            precompLayerRegistry[lIdx] = shotName; // store first shot name
                        }
                        var rangeBinName = precompLayerRegistry[lIdx] + "_" + pad(currentNum, 3);
                        if (!precompRangeBinRegistry[lIdx]) {
                            precompRangeBinRegistry[lIdx] = getShotBin(binShots, rangeBinName);
                        } else {
                            try { precompRangeBinRegistry[lIdx].name = rangeBinName; } catch(eBin) {}
                        }
                        try {
                            var srcItem = mainComp.layer(lIdx).source;
                            srcItem.name         = rangeBinName + "_container" + osSuffix;
                            srcItem.parentFolder = precompRangeBinRegistry[lIdx];
                        } catch(eRn2) {}
                    }

                    // No container comp. Mark source as processed, replace footage layer
                    // inside the deepest precomp with shotComp. All transforms/effects on
                    // footageLayer are preserved in place — nothing to transplant.
                    // Key on the ORIGINAL source id (pre-replaceSource) so later iterations
                    // can still detect the dupe after this replacement mutates the live source.
                    if (item.originalSourceId) processedSourceIds[item.originalSourceId] = { bin: shotBin };

                    plateInner.property("Marker").setValueAtTime(cutStart, cutMarker("cut in"));
                    plateInner.property("Marker").setValueAtTime(cutStart + cutDuration, cutMarker("cut out"));

                    footageLayer.replaceSource(shotComp, false);

                    // When overscan is active the shotComp is larger than the original footage.
                    // The footageLayer's anchor was set for the original footage size, so sampling
                    // is offset by half the overscan margin. Shift the anchor to compensate.
                    if (shotOverscan > 0) {
                        try {
                            var osOffX = (osWidth  - source.width)  / 2;
                            var osOffY = (osHeight - source.height) / 2;
                            var apFL = footageLayer.property("ADBE Transform Group").property("ADBE Anchor Point");
                            var curFL = apFL.value;
                            apFL.setValue([curFL[0] + osOffX, curFL[1] + osOffY]);
                        } catch(eOsAP) {}
                    }

                    // Trim the footageLayer (now showing shotComp) to the rendered plate range.
                    // Also capture the cut/plate times in footageComp time — these are used below
                    // to drive both the main-comp layer trim and the _dynamicLink comp, so that
                    // untrimmed precomp layers (outPoint at comp end) don't produce wrong results.
                    var flPlateStart = 0, flPlateEnd = 0, flCutIn = 0, flCutOut = 0;
                    try {
                        var flStretch = (footageLayer.stretch !== 0) ? footageLayer.stretch : 100;
                        if (footageLayer.timeRemapEnabled) {
                            // Source→comp mapping is nonlinear for time-remapped layers.
                            // The footage layer's in/out are already correct cut bounds in comp time.
                            // Estimate handle offsets using the local comp/source time ratio near the cut.
                            var trCompDur = footageLayer.outPoint - footageLayer.inPoint;
                            var trSrcDur  = (range.end > range.start) ? (range.end - range.start) : 1;
                            // Guard both dividend and divisor — a zero-duration footage layer
                            // (manually trimmed to a point) produces trCompDur === 0 which
                            // cascades to negative/zero handle offsets below.
                            var trRate    = (trCompDur > 0) ? (trCompDur / trSrcDur) : 1;
                            flCutIn      = footageLayer.inPoint;
                            flCutOut     = footageLayer.outPoint;
                            flPlateStart = flCutIn  - handleSec * trRate;
                            flPlateEnd   = flCutOut + handleSec * trRate;
                        } else {
                            flPlateStart = footageLayer.startTime + fullStart * (flStretch / 100);
                            flPlateEnd   = footageLayer.startTime + (fullStart + fullDurationSec) * (flStretch / 100);
                            flCutIn      = footageLayer.startTime + cutStart * (flStretch / 100);
                            flCutOut     = footageLayer.startTime + (cutStart + cutDuration) * (flStretch / 100);
                        }
                        footageLayer.inPoint  = flPlateStart;
                        footageLayer.outPoint = flPlateEnd;
                        footageLayer.property("Marker").setValueAtTime(flCutIn,  cutMarker("cut in"));
                        footageLayer.property("Marker").setValueAtTime(flCutOut, cutMarker("cut out"));
                    } catch(eTrim) {}

                    var precompSrc = layer.source;
                    var blPC = shotComp.layers.byName("GUIDE_BURNIN"); if (blPC) { blPC.locked=false; blPC.moveToBeginning(); blPC.locked=true; }
                    containerDurFrames = Math.round(precompSrc.duration * safeFPS);

                } else {
                    // ── 2. Container (direct footage path) ────────────────────
                    // Native precompose with moveAllAttributes=true. AE handles
                    // effect stacking order, masks, transforms, time remap,
                    // expressions, layer flags (motion blur, 3D, quality, frame
                    // blending, sampling, preserve-transparency), blending mode,
                    // parenting, and file bindings on effects like Apply Color
                    // LUT — all natively and without the silent correctness
                    // bugs that our hand-rolled moveLayerAttribs had
                    // (effect-order reversal, dropped layer flags, LUT locate
                    // dialogs). The precomp path (isPrecomp branch above)
                    // already uses this same pattern; we now use it here too
                    // for the direct-footage path.
                    //
                    // Capture the ORIGINAL layer's timing before precompose
                    // invalidates the reference. AE's native precompose does NOT
                    // reliably preserve the mainComp precomp layer's in/out —
                    // the new precomp layer tends to span the full comp by
                    // default, which would wipe out the edit's cut placement.
                    // We restore the original in/out/startTime ourselves below.
                    var origInPoint   = layer.inPoint;
                    var origOutPoint  = layer.outPoint;
                    var layerIdx      = layer.index; // read before precompose invalidates the ref
                    var containerComp;
                    try {
                        containerComp = mainComp.layers.precompose([layerIdx], shotName + "_container" + osSuffix, true);
                    } catch (ePC) {
                        reportError("PREP", ePC, "Direct-footage precompose failed for " + shotName);
                        return;
                    }
                    containerComp.displayStartTime = 0;
                    containerComp.parentFolder     = shotBin;
                    containerComp.label            = 8; // Blue — container (editorial)

                    // The original footage layer now lives inside containerComp with
                    // every attribute moved in by AE. Find it, then swap its source
                    // to the shot's render target (shotComp).
                    var containerInner = null;
                    for (var ci = 1; ci <= containerComp.numLayers; ci++) {
                        var cl = containerComp.layer(ci);
                        try { if (cl.source === source) { containerInner = cl; break; } } catch(eCL) {}
                    }
                    if (!containerInner && containerComp.numLayers > 0) {
                        // Fallback: precompose only put one layer in if we got here.
                        containerInner = containerComp.layer(1);
                    }

                    // Shrink containerComp from mainComp.duration down to the
                    // shot's own cut+handles length so opening the container
                    // shows a clean 0-based timeline instead of a huge amount
                    // of pre/post dead space.
                    try {
                        containerComp.duration = fullDurationSec;
                        containerComp.displayStartTime = 0;
                    } catch(eDur) {}

                    if (containerInner) {
                        containerInner.replaceSource(shotComp, false);
                        // Shift containerInner so the first rendered frame sits
                        // at container time 0 (edit-friendly). shotComp's rendered
                        // content lives at shotComp-time fullStart..fullEnd;
                        // startTime = -fullStart makes container time 0 line up
                        // with that first rendered frame.
                        containerInner.startTime = -fullStart;
                        containerInner.inPoint   = 0;
                        containerInner.outPoint  = fullDurationSec;
                        // Markers now in container-local time: cut_in at handleSec,
                        // cut_out at handleSec + cutDuration. Same for
                        // containerInner (layer-level). plateInner markers stay at
                        // source-time cutStart/cutEnd because shotComp's internal
                        // timeline IS source time.
                        containerComp.markerProperty.setValueAtTime(handleSec,               cutMarker("cut in"));
                        containerComp.markerProperty.setValueAtTime(handleSec + cutDuration, cutMarker("cut out"));
                        plateInner.property("Marker").setValueAtTime(cutStart,               cutMarker("cut in"));
                        plateInner.property("Marker").setValueAtTime(cutStart + cutDuration, cutMarker("cut out"));
                        containerInner.property("Marker").setValueAtTime(handleSec,               cutMarker("cut in"));
                        containerInner.property("Marker").setValueAtTime(handleSec + cutDuration, cutMarker("cut out"));
                        var blDirect = shotComp.layers.byName("GUIDE_BURNIN");
                        if (blDirect) { blDirect.locked = false; blDirect.moveToBeginning(); blDirect.locked = true; }

                        // Overscan: containerInner's anchor was set for source-sized
                        // footage, but shotComp is osWidth × osHeight. Shift to
                        // re-center sampling on the inflated source.
                        if (shotOverscan > 0) {
                            try {
                                var osOffX = (osWidth  - source.width)  / 2;
                                var osOffY = (osHeight - source.height) / 2;
                                var apCI = containerInner.property("ADBE Transform Group").property("ADBE Anchor Point");
                                var curCI = apCI.value;
                                apCI.setValue([curCI[0] + osOffX, curCI[1] + osOffY]);
                            } catch(eOsAP2) {}
                        }
                    }

                    // Re-bind `layer` to the new precomp layer in mainComp so the
                    // downstream `layer.inPoint` read for Nuke data still works.
                    // AE places the new precomp layer at the same index the original
                    // occupied; if anything shifted, scan for the layer whose source
                    // is our new containerComp.
                    try { layer = mainComp.layer(layerIdx); } catch(eLIdx) {}
                    var sourceMatchOK = false;
                    try { sourceMatchOK = (layer && layer.source === containerComp); } catch(eSM) {}
                    if (!sourceMatchOK) {
                        for (var mL = 1; mL <= mainComp.numLayers; mL++) {
                            try {
                                if (mainComp.layer(mL).source === containerComp) { layer = mainComp.layer(mL); break; }
                            } catch(eML) {}
                        }
                    }

                    // Align the mainComp precomp layer with the now-shifted
                    // container timeline. Container time 0 = first rendered handle
                    // frame; we want that to appear at mainComp time
                    // (origInPoint - handleSec) so the cut still lands exactly at
                    // origInPoint..origOutPoint (the editorial cut placement).
                    if (layer) {
                        try {
                            layer.startTime = origInPoint - handleSec;
                            layer.inPoint   = origInPoint;
                            layer.outPoint  = origOutPoint;
                        } catch (eTiming) {}
                    }

                    // Patch every selLayers / expandedLayers entry that still holds a
                    // reference to the pre-precompose layer. Without this, the
                    // dynamicLink build loop later iterates selLayers and hits stale
                    // "Object is invalid" refs for every layer that went through
                    // native precompose.
                    for (var sUpd = 0; sUpd < selLayers.length; sUpd++) {
                        if (selLayers[sUpd].mainLayerIdx === layerIdx) {
                            selLayers[sUpd].layer     = layer;
                            selLayers[sUpd].isPrecomp = true;
                        }
                    }
                    for (var eUpd = i + 1; eUpd < expandedLayers.length; eUpd++) {
                        if (expandedLayers[eUpd].mainLayerIdx === layerIdx) {
                            expandedLayers[eUpd].layer = layer;
                        }
                    }

                    containerDurFrames = Math.round(source.duration * safeFPS);
                }

                // ── Render queue ───────────────────────────────────────────────
                var rq = proj.renderQueue.items.add(shotComp);
                // Epsilon pushes timeSpanStart past any frame boundary, never reaching the buffer frame
                rq.timeSpanStart    = fullStart + 0.0001;
                rq.timeSpanDuration = fullDurationSec;

                var om = rq.outputModule(1);
                var foundT = false;
                for (var t = 0; t < om.templates.length; t++) if (om.templates[t] === omTemplate) foundT = true;
                if (foundT) om.applyTemplate(omTemplate);

                var fsShotDir = new Folder(fsScripts.fsName + "/" + shotName);
                if (!fsShotDir.exists) fsShotDir.create();
                var fsShotPlate = new Folder(fsShotDir.fsName + "/plate"); if(!fsShotPlate.exists) fsShotPlate.create();
                var fsShotRender = new Folder(fsShotDir.fsName + "/render"); if(!fsShotRender.exists) fsShotRender.create();
                var pPath = new File(fsShotPlate.fsName + "/" + shotName + "_plate.mov");
                om.file = pPath;

                renderItems.push({ n: shotName, p: pPath, c: shotComp, s: fullStart, w: osWidth, h: osHeight, cs: cutStart, cd: cutDuration, bin: shotBin });

                var cutDurationFrames = Math.round(cutDuration * safeFPS);
                nukeDataList.push({
                    name:                    shotName,
                    platePath:               pPath.fsName,
                    renderPath:              fsShotRender.fsName + "/" + shotName + "_render_v01.mov",
                    globalStartTime:         layer.inPoint,
                    handles:                 actualLeadingHandles,
                    fullDurationFrames:      fullDurationFrames,
                    cutDurationFrames:       cutDurationFrames,

                    containerDurationFrames: containerDurFrames,
                    cutStartFrames:          Math.round(cutStart * safeFPS),
                    w:                       osWidth,
                    h:                       osHeight,
                    origW:                   source.width,
                    origH:                   source.height
                });
            }

        } catch(e) { reportError("PREP", e); try { progress.close(); } catch(eCP){} return; } finally { app.endUndoGroup(); }

        if (cancelCheck()) return;
        if (!chkSkipRender.value) {
            // Close the palette for the duration of the render: AE blocks the
            // script during renderQueue.render() so palette updates freeze
            // anyway, and AE's own render window takes over the UI — use its
            // cancel button if you need to stop the render mid-flight.
            progress.update("Saving project and handing off to AE's render queue\u2026", "AE's render window will take over now.", 60);
            progress.close();
            try { proj.save(); proj.renderQueue.render(); } catch(e) { reportError("RENDER", e, "Check OM Template and Disk Space."); return; }
            progress = makeProgressPanel();
            progress.update("Render finished, importing plates\u2026", "", 65);
        }

        app.beginUndoGroup("Roundtrip Finish");
        try {
            var count = 0;
            var importedShots = [];
            var missingPlates = [];
            for(var k=0; k<renderItems.length; k++) {
                if (cancelCheck()) return;
                progress.update(null,
                    "importing plate " + (k + 1) + " of " + renderItems.length + ": " + renderItems[k].n,
                    65 + 15 * (k / Math.max(1, renderItems.length)));
                var ri = renderItems[k];
                var readyToImport = false;
                if (chkSkipRender.value) {
                    // Debug: plate was never rendered — import directly from shots folder if it exists
                    readyToImport = ri.p.exists;
                } else {
                    if (!waitForFile(ri.p, 20)) {
                        missingPlates.push(ri.n + " (plate not found — disk full or render failed?)");
                        alert("Plate not found after render:\n" + ri.p.fsName + "\n\nDisk full or render failed?");
                    } else {
                        readyToImport = true;
                    }
                }
                if (readyToImport) {
                    var imp = proj.importFile(new ImportOptions(ri.p));
                    imp.parentFolder = ri.bin;
                    imp.label = 9; // Green — VFX return mov
                    var tComp = ri.c;
                    if(tComp instanceof CompItem) {
                        var nL = tComp.layers.add(imp);
                        nL.startTime = ri.s;
                        nL.position.setValue([ri.w/2, ri.h/2]);
                        nL.label = 11;
                        nL.property("Marker").setValueAtTime(ri.cs, cutMarker("cut in"));
                        nL.property("Marker").setValueAtTime(ri.cs + ri.cd, cutMarker("cut out"));
                        for(var xx=1; xx<=tComp.numLayers; xx++) {
                            var L = tComp.layer(xx);
                            if(L!==nL && !L.guideLayer) { L.enabled=false; L.audioEnabled=false; }
                        }
                        var bl=tComp.layers.byName("GUIDE_BURNIN"); if(bl) { bl.locked=false; bl.moveToBeginning(); bl.locked=true; }
                        importedShots.push(ri.n);
                        count++;
                    }
                }
            }
            stats.count = count;

            // ── dynamicLink Build ──────────────────────────────────────────────
            // For every selected layer whose current source is a CompItem
            // (precomps natively + direct footage that the main loop has
            // already wrapped into a _container), build a <source>_dynamicLink
            // wrapper comp in /Shots/dynamicLink with duration exactly
            // cut + 2×handleFrames. Uses the proven extend → wrap → contract
            // sequence so time-remap / time-stretch layers get correct source
            // offsets. Any handle time lost to comp-edge clamping is padded
            // with black (source out of range).
            var dynBuilt       = 0;
            var dynBuildFails  = [];
            var dynBuildSkips  = []; // verbose diagnostics for non-processed entries
            if (chkCreateDynLink.value && selLayers && selLayers.length > 0) {
                progress.update("Building Dynamic Link wrappers\u2026", "0 of " + selLayers.length, 80);
                var binDynLink = getShotBin(getBinFolder("Shots"), "dynamicLink");
                var handleSec  = handleFrames / mainComp.frameRate;

                for (var dli = 0; dli < selLayers.length; dli++) {
                    if (cancelCheck()) return;
                    // ── readable label for diagnostics (try/catch because the
                    //    layer DOM may be invalidated by prior processing) ──
                    var dlTag = "selLayers[" + dli + "]";
                    try { if (selLayers[dli] && selLayers[dli].layer) dlTag = selLayers[dli].layer.name; } catch(eTag) {}
                    progress.update(null,
                        (dli + 1) + " of " + selLayers.length + ": " + dlTag,
                        80 + 15 * (dli / Math.max(1, selLayers.length)));

                    var dlLayer = selLayers[dli].layer;
                    if (!dlLayer) {
                        dynBuildSkips.push(dlTag + ": layer reference is null");
                        continue;
                    }
                    // Process the layer if its CURRENT source is a comp.
                    // - Precomp selections: source has always been a comp.
                    // - Direct-footage selections: the main loop's container
                    //   wrapping + replaceSource() has already swapped the
                    //   source to a CompItem by the time we get here.
                    // Anything still pointing at raw footage was never
                    // processed by the main loop and can't be wrapped here.
                    try {
                        if (!(dlLayer.source instanceof CompItem)) {
                            dynBuildSkips.push(dlTag + ": source is not a CompItem (unprocessed footage?)");
                            continue;
                        }
                    } catch (eValid) {
                        dynBuildSkips.push(dlTag + ": layer object invalidated — " + eValid.toString());
                        continue;
                    }

                    try {
                        var origIn  = dlLayer.inPoint;
                        var origOut = dlLayer.outPoint;
                        var cutSec  = origOut - origIn;

                        // Step 1 — extend (captured-target pattern, matches extend_layer_trim)
                        var extTargetOut = Math.min(mainComp.duration, origOut + handleSec);
                        dlLayer.inPoint  = Math.max(0, origIn - handleSec);
                        dlLayer.outPoint = extTargetOut;

                        // Measure what we actually achieved (may be clipped at comp edges)
                        var actualLead  = origIn - dlLayer.inPoint;   // >= 0
                        var leadLostSec = handleSec - actualLead;      // black pad at head

                        // Step 2 — create dynamicLink comp
                        var dlSrc    = dlLayer.source;
                        var srcAtIn  = mapTimeToSource(dlLayer, dlLayer.inPoint);
                        var dlDurSec = cutSec + 2 * handleSec; // always full length

                        var dlComp = proj.items.addComp(
                            dlSrc.name + "_dynamicLink",
                            dlSrc.width, dlSrc.height, dlSrc.pixelAspect,
                            dlDurSec, dlSrc.frameRate
                        );
                        dlComp.displayStartTime = 0;
                        dlComp.parentFolder     = binDynLink;
                        dlComp.label            = 14; // Cyan

                        var dlInner = dlComp.layers.add(dlSrc);
                        // At dynCompTime = handleSec the inner layer shows the cut-start frame.
                        //   cut_start_in_dynamicLink = leadLostSec + actualLead = handleSec ✓
                        //   inner shows source at (dynCompTime - dlInner.startTime)
                        dlInner.startTime = leadLostSec - srcAtIn;

                        // Cut markers — always at handleSec / handleSec+cutSec by construction
                        dlComp.markerProperty.setValueAtTime(handleSec,          cutMarker("cut in"));
                        dlComp.markerProperty.setValueAtTime(handleSec + cutSec, cutMarker("cut out"));

                        // Step 3 — contract back (captured-target, matches contract_layer_trim:
                        // set inPoint first so AE's duration preservation doesn't drag outPoint)
                        var contractTarget = origOut;
                        dlLayer.inPoint  = origIn;
                        dlLayer.outPoint = contractTarget;

                        dynBuilt++;
                    } catch (eDL) {
                        var dlName = "layer " + dli;
                        try { dlName = dlLayer.name; } catch(eName) {}
                        dynBuildFails.push(dlName + ": " + eDL.toString());
                        // Best-effort rollback so the main comp layer doesn't stay extended
                        try { dlLayer.inPoint = origIn; dlLayer.outPoint = origOut; } catch(eRb) {}
                    }
                }
            }

            if (nukeDataList.length > 0) {
                if (chkCreateNuke.value) {
                    var nukeFileName = proj.file.name.replace(".aep", "") + "_Comp.nk";
                    var nukeFile = new File(fsScripts.fsName + "/" + nukeFileName);
                    writeNukeScript(nukeFile, nukeDataList, stats.fps);

                    for (var nsi = 0; nsi < nukeDataList.length; nsi++) {
                        var nsd = nukeDataList[nsi];
                        var shotNkFile = new File(fsScripts.fsName + "/" + nsd.name + "/" + nsd.name + ".nk");
                        writeNukeShotScript(shotNkFile, nsd, stats.fps);
                    }
                }

            }

            if (chkExportXML.value) {
                var xmlScript = new File((new File($.fileName)).parent.parent.fsName + "/export-shot-xml/export_shot_xml.jsx");
                if (xmlScript.exists) {
                    $.global.__shotRoundtripXMLDir = fsShots.fsName;
                    $.evalFile(xmlScript);
                    delete $.global.__shotRoundtripXMLDir;
                } else {
                    alert("Export Shot XML script not found:\n" + xmlScript.fsName);
                }
            }

            try { proj.save(); } catch(eSave) {}

            var rpt = new Window("dialog", "Roundtrip Complete");
            rpt.orientation = "column"; rpt.alignChildren = ["fill", "top"];
            rpt.spacing = 10; rpt.margins = 14;

            var statRow = function(parent, label, value) {
                var g = parent.add("group");
                g.orientation = "row"; g.alignChildren = ["left", "center"]; g.spacing = 8;
                var l = g.add("statictext", undefined, label); l.preferredSize = [110, 20];
                g.add("statictext", undefined, value);
            };

            var pStats = rpt.add("panel", undefined, "Results");
            pStats.orientation = "column"; pStats.alignChildren = ["fill", "top"];
            pStats.spacing = 5; pStats.margins = [10, 15, 10, 10];

            statRow(pStats, "Shots processed:", chkSkipRender.value ? "— (render skipped)" : "" + stats.count);
            var resTxt = stats.origWidth + " \u2192 " + stats.plateWidth + " \u00d7 " + stats.plateHeight + " px";
            if (overscanPercent > 0) {
                var osOnCount = 0;
                for (var ori = 0; ori < expandedLayers.length; ori++) { if (expandedLayers[ori].overscan) osOnCount++; }
                if (osOnCount === expandedLayers.length) resTxt += "  (+" + overscanPercent + "% overscan)";
                else if (osOnCount > 0) resTxt += "  (+" + overscanPercent + "% overscan on " + osOnCount + "/" + expandedLayers.length + " shots)";
            }
            statRow(pStats, "Resolution:", resTxt);
            statRow(pStats, "Start frame:", "1001 (cut)");
            statRow(pStats, "Handles:", handleFrames + " frames");
            statRow(pStats, "dynamicLink built:", "" + dynBuilt);

            if (importedShots.length > 0) {
                var pImport = rpt.add("panel", undefined, "\u2713 Imported (" + importedShots.length + ")");
                pImport.orientation = "column"; pImport.alignChildren = ["fill", "top"];
                pImport.spacing = 3; pImport.margins = [10, 15, 10, 10];
                var maxI = Math.min(importedShots.length, 10);
                for (var ii = 0; ii < maxI; ii++) { pImport.add("statictext", undefined, importedShots[ii]); }
                if (importedShots.length > 10) pImport.add("statictext", undefined, "...and " + (importedShots.length - 10) + " more.");
            }

            if (missingPlates.length > 0) {
                var pMiss = rpt.add("panel", undefined, "! Missing Plates (" + missingPlates.length + ")");
                pMiss.orientation = "column"; pMiss.alignChildren = ["fill", "top"];
                pMiss.spacing = 3; pMiss.margins = [10, 15, 10, 10];
                for (var mi = 0; mi < missingPlates.length; mi++) { pMiss.add("statictext", undefined, missingPlates[mi]); }
            }

            if (clampedShots.length > 0) {
                var pClamp = rpt.add("panel", undefined, "! Handles Clamped");
                pClamp.orientation = "column"; pClamp.alignChildren = ["fill", "top"];
                pClamp.spacing = 4; pClamp.margins = [10, 15, 10, 10];
                var maxC = Math.min(clampedShots.length, 8);
                for (var c = 0; c < maxC; c++) { pClamp.add("statictext", undefined, clampedShots[c]); }
                if (clampedShots.length > 8) pClamp.add("statictext", undefined, "...and " + (clampedShots.length - 8) + " more.");
            }

            if (skippedLayers.length > 0) {
                var pWarn = rpt.add("panel", undefined, "! Ignored Layers");
                pWarn.orientation = "column"; pWarn.alignChildren = ["fill", "top"];
                pWarn.spacing = 4; pWarn.margins = [10, 15, 10, 10];
                var maxS = Math.min(skippedLayers.length, 5);
                for (var s = 0; s < maxS; s++) { pWarn.add("statictext", undefined, skippedLayers[s]); }
                if (skippedLayers.length > 5) pWarn.add("statictext", undefined, "...and " + (skippedLayers.length - 5) + " more.");
            }

            if (dynBuildFails.length > 0) {
                var pDlErr = rpt.add("panel", undefined, "! dynamicLink Errors (" + dynBuildFails.length + ")");
                pDlErr.orientation = "column"; pDlErr.alignChildren = ["fill", "top"];
                pDlErr.spacing = 4; pDlErr.margins = [10, 15, 10, 10];
                var maxDE = Math.min(dynBuildFails.length, 8);
                for (var de = 0; de < maxDE; de++) { pDlErr.add("statictext", undefined, dynBuildFails[de]); }
                if (dynBuildFails.length > 8) pDlErr.add("statictext", undefined, "...and " + (dynBuildFails.length - 8) + " more.");
            }

            if (dynBuildSkips.length > 0) {
                var pDlSk = rpt.add("panel", undefined, "dynamicLink Skipped (" + dynBuildSkips.length + ")");
                pDlSk.orientation = "column"; pDlSk.alignChildren = ["fill", "top"];
                pDlSk.spacing = 4; pDlSk.margins = [10, 15, 10, 10];
                var maxDS = Math.min(dynBuildSkips.length, 8);
                for (var ds = 0; ds < maxDS; ds++) { pDlSk.add("statictext", undefined, dynBuildSkips[ds]); }
                if (dynBuildSkips.length > 8) pDlSk.add("statictext", undefined, "...and " + (dynBuildSkips.length - 8) + " more.");
            }

            var btnGrpR = rpt.add("group");
            btnGrpR.orientation = "row"; btnGrpR.alignment = ["fill", "bottom"];
            var brandLbl = btnGrpR.add("statictext", undefined, "Gegenschuss \u00b7 AE Shot Roundtrip");
            brandLbl.alignment = ["left", "center"];
            var spacerR = btnGrpR.add("statictext", undefined, ""); spacerR.alignment = ["fill", "center"];
            var btnClose = btnGrpR.add("button", undefined, "Close"); btnClose.preferredSize = [80, 28];
            btnClose.onClick = function() { rpt.close(); };
            progress.update("Done.", "", 100);
            progress.close();
            rpt.show();

        } catch(e) { reportError("FINISH", e); try { progress.close(); } catch(eCF){} } finally { app.endUndoGroup(); }
    }
    vfxRoundtripEpsilon();
}