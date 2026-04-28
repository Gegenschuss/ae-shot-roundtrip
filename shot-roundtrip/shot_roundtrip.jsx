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

  4.  Queues  shotname_comp  in the AE render queue. Output: ProRes 422 HQ .mov
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
  OM Template     AE Output Module template name for the render (e.g. ProRes 422 HQ).
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
        _grade/             ← Resolve graded returns (flat, shared)
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

        var dlg = new Window("dialog", "Shot Roundtrip");
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
        var chkAutoStart = r2.add("checkbox", undefined, "Auto");
        chkAutoStart.helpTip = "Auto-pick shot numbers from the selected layers' positions in the main comp. "
                             + "Sandwiches new shots between existing ones (e.g. shot_035 between _030 and _040).";
        function applyAutoStartUI() {
            etStartNum.enabled = !chkAutoStart.value;
        }
        chkAutoStart.onClick = applyAutoStartUI;

        var r2b = addRow(pnlMain, "Increment:");
        var etIncrement = r2b.add("edittext", undefined, "10");
        etIncrement.preferredSize = [60, FIELD_H];

        var r3 = addRow(pnlMain, "Handles:");
        var etHandles = r3.add("edittext", undefined, "50");
        etHandles.preferredSize = [60, FIELD_H];
        r3.add("statictext", undefined, "frames");

        // Color helper: paint the active comp's top-level layers sandstone,
        // override REVERSED layers (negative stretch OR descending time-
        // remap) to blue. Uses AE label slots 15 (Sandstone) and 8 (Blue) —
        // slots are user-customisable in Preferences, same caveat as the
        // rest of AE's label system. Detection mirrors the reversal scan
        // (describeTimeEffect → .reversed), so what gets blued here is
        // exactly what the warning dialog later flags.
        var rColor = pnlMain.add("group");
        rColor.orientation = "row"; rColor.alignChildren = ["left", "center"]; rColor.spacing = 8;
        var rColorPad = rColor.add("statictext", undefined, ""); rColorPad.preferredSize = [LABEL_W, FIELD_H];
        var btnColorTR = rColor.add("button", undefined, "Color Time-Reverse Layers");
        btnColorTR.preferredSize = [220, FIELD_H + 2];
        btnColorTR.helpTip = "Walk the active comp's top-level layers: paint everything sandstone, then paint reversed layers (negative stretch or descending time-remap) blue. Wraps in one undo step.";
        var LABEL_SAND = 15;
        var LABEL_BLUE = 8;
        btnColorTR.onClick = function() {
            try {
                // mainComp isn't assigned until later in the script flow,
                // so resolve via proj.activeItem here. The pre-dialog check
                // already guarantees activeItem is a CompItem.
                var comp = proj && proj.activeItem;
                if (!comp || !(comp instanceof CompItem) || !comp.numLayers) {
                    alert("No comp layers to color.");
                    return;
                }
                app.beginUndoGroup("Color Time-Reverse Layers");
                try {
                    var painted = 0, hadReverse = 0;
                    for (var li = 1; li <= comp.numLayers; li++) {
                        var L = comp.layer(li);
                        try {
                            // describeTimeEffect returns null for layers
                            // with no time effect; { reversed: bool, ... }
                            // otherwise. Hoisted from elsewhere in this
                            // function, so it's available here.
                            var fx = null;
                            try { fx = describeTimeEffect(L); } catch (eFx) {}
                            var isRev = !!(fx && fx.reversed);
                            L.label = isRev ? LABEL_BLUE : LABEL_SAND;
                            painted++;
                            if (isRev) hadReverse++;
                        } catch (eL) {}
                    }
                    alert("Colored " + painted + " layer" + (painted === 1 ? "" : "s") +
                          " — " + hadReverse + " reversed (blue), the rest sandstone.");
                } finally {
                    app.endUndoGroup();
                }
            } catch (eC) {
                alert("Color Time-Reverse Layers failed:\n" + eC.message);
            }
        };

        // ── Pipeline Options ───────────────────────────
        var pnlOpt = dlg.add("panel", undefined, "Pipeline Options");
        pnlOpt.orientation = "column"; pnlOpt.alignChildren = ["fill", "top"];
        pnlOpt.spacing = 6; pnlOpt.margins = [10, 15, 10, 10];

        var chkCreateNuke     = pnlOpt.add("checkbox", undefined, "Create Nuke Scripts");  chkCreateNuke.value     = true;
        var chkExportXML      = pnlOpt.add("checkbox", undefined, "Export Shot XML");    chkExportXML.value      = true;
        var chkCreateDynLink  = pnlOpt.add("checkbox", undefined, "Create dynamicLink Comps"); chkCreateDynLink.value = true;
        // Note: duplicate-shared-sources is a contextual choice — it only
        // matters when the selection actually contains shared inner
        // footage. Its checkbox lives in the Confirm Shots preflight
        // dialog (alongside the shared-source warning panel) rather
        // than here, so the user only sees it when it's relevant. The
        // persisted default still travels through this SR_DEFAULTS /
        // srLoad chain so the last-used value survives across runs.

        var r4 = addRow(pnlOpt, "Overscan:");
        var etOverscan = r4.add("edittext", undefined, "10");
        etOverscan.preferredSize = [50, FIELD_H];
        r4.add("statictext", undefined, "%");

        var r5 = addRow(pnlOpt, "OM Template:");
        var etOM = r5.add("edittext", undefined, "ProRes 422 HQ");
        etOM.preferredSize = [150, FIELD_H];

        var r6 = addRow(pnlOpt, "Shots Folder:");
        var etShotsFolder = r6.add("edittext", undefined, "../Roundtrip");
        etShotsFolder.preferredSize = [150, FIELD_H];
        var btnBrowseShots = r6.add("button", undefined, "Browse\u2026");
        btnBrowseShots.preferredSize = [80, FIELD_H];
        btnBrowseShots.onClick = function () {
            // Seed the picker with the current field's resolved path when it
            // exists, so the user doesn't start from ~/ every time.
            var seed = null;
            try {
                var txt = etShotsFolder.text || "";
                var candidate = /^(\/|[A-Za-z]:)/.test(txt)
                              ? new Folder(txt)
                              : (proj.file ? new Folder(proj.file.parent.fsName + "/" + txt) : null);
                if (candidate && candidate.exists) seed = candidate;
            } catch (eSeed) {}
            var picked = (seed ? seed : Folder.desktop).selectDlg("Select Shots Folder");
            if (picked) etShotsFolder.text = picked.fsName;
        };

        var chkSkipRender     = pnlOpt.add("checkbox", undefined, "Skip Render (debug)");           chkSkipRender.value     = false;

        // ── Burnin Fields ──────────────────────────────
        // Populate the "Burnin Fields" precomp at roundtrip time so the
        // burnin prefix is ready before the user opens the first shot.
        // Pre-populated from the existing Burnin Fields comp (if present)
        // so re-runs don't clobber the user's edits; otherwise the defaults
        // below (Gegenschuss as the production company, .aep filename as
        // project) kick in.
        var pnlBurnin = dlg.add("panel", undefined, "Burnin");
        pnlBurnin.orientation = "column"; pnlBurnin.alignChildren = ["fill", "top"];
        pnlBurnin.spacing = 6; pnlBurnin.margins = [10, 15, 10, 10];

        // Master "Include Burnin" checkbox at the top of the panel.
        // When unchecked, the roundtrip skips creating both the Burnin
        // Fields precomp and the Burnin overlay on mainComp. The field
        // rows stay visible regardless so the user can prep values
        // without toggling the master switch first.
        var gBurninOn = pnlBurnin.add("group");
        gBurninOn.orientation = "row"; gBurninOn.alignChildren = ["left", "center"]; gBurninOn.spacing = 8;
        gBurninOn.add("statictext", undefined, "").preferredSize = [140, FIELD_H];
        var chkBurninEnabled = gBurninOn.add("checkbox", undefined, "Include Burnin");
        chkBurninEnabled.value = true;

        function addBurninRow(parent, labelText) {
            var g = parent.add("group");
            g.orientation = "row"; g.alignChildren = ["left", "center"]; g.spacing = 8;
            var lbl = g.add("statictext", undefined, labelText);
            lbl.preferredSize = [140, FIELD_H];
            var et = g.add("edittext", undefined, "");
            et.preferredSize = [220, FIELD_H];
            return et;
        }
        var etBurninProject = addBurninRow(pnlBurnin, "Project:");
        var etBurninCompany = addBurninRow(pnlBurnin, "Production Company:");
        var etBurninAgency  = addBurninRow(pnlBurnin, "Agency:");
        var etBurninClient  = addBurninRow(pnlBurnin, "Client:");



        // ── Settings persistence ───────────────────────
        // All fields except chkSkipRender (debug) round-trip through
        // app.settings so the dialog remembers last-used values. Reset to
        // Defaults populates the hardcoded defaults in-place; changes only
        // persist when the user clicks Run Roundtrip.
        var SR_SECTION  = "Gegenschuss Shot Roundtrip";
        var SR_DEFAULTS = {
            prefix:        "shot_",
            startNum:      "010",
            autoStart:     "true",
            increment:     "10",
            handles:       "50",
            createNuke:        "true",
            exportXML:         "true",
            createDynLink:     "true",
            sharedSourceMode:  "separate",
            overscan:      "10",
            omTemplate:    "ProRes 422 HQ",
            shotsFolder:   "/Roundtrip",
            // Burnin fields — empty project falls back to the .aep filename
            // stem at apply-time (see srApply). Company seeded to "Gegenschuss".
            burninEnabled:  "true",
            burninProject:  "",
            burninCompany:  "Gegenschuss",
            burninAgency:   "",
            burninClient:   ""
        };
        function srLoad(key, fallback) {
            try {
                if (app.settings.haveSetting(SR_SECTION, key)) return app.settings.getSetting(SR_SECTION, key);
            } catch (e) {}
            return fallback;
        }
        function srSave(key, value) {
            try { app.settings.saveSetting(SR_SECTION, key, String(value)); } catch (e) {}
        }
        function srApply(s) {
            etPrefix.text          = s.prefix;
            etStartNum.text        = s.startNum;
            chkAutoStart.value     = (s.autoStart    === "true" || s.autoStart    === true);
            etIncrement.text       = s.increment;
            etHandles.text         = s.handles;
            chkCreateNuke.value    = (s.createNuke    === "true" || s.createNuke    === true);
            chkExportXML.value     = (s.exportXML     === "true" || s.exportXML     === true);
            chkCreateDynLink.value = (s.createDynLink === "true" || s.createDynLink === true);
            // sharedSourceMode has no main-dialog control — its dedicated
            // preflight dialog reads srLoad("sharedSourceMode", …) at
            // show-time and saves the user's pick on Continue.
            etOverscan.text        = s.overscan;
            etOM.text              = s.omTemplate;
            etShotsFolder.text     = s.shotsFolder;
            // Burnin fields — empty project falls back to the .aep stem so
            // Reset to Defaults gets the current project's name instead of
            // blanking the field.
            var projStemForBurn = (app.project.file && app.project.file.displayName)
                                ? app.project.file.displayName.replace(/\.aep$/i, "")
                                : "Project";
            if ("burninEnabled" in s) chkBurninEnabled.value = (s.burninEnabled === "true" || s.burninEnabled === true);
            if ("burninProject" in s) etBurninProject.text = s.burninProject && s.burninProject.length > 0 ? s.burninProject : projStemForBurn;
            if ("burninCompany" in s) etBurninCompany.text = s.burninCompany;
            if ("burninAgency"  in s) etBurninAgency.text  = s.burninAgency;
            if ("burninClient"  in s) etBurninClient.text  = s.burninClient;
            applyAutoStartUI();
        }
        srApply({
            prefix:         srLoad("prefix",         SR_DEFAULTS.prefix),
            startNum:       srLoad("startNum",       SR_DEFAULTS.startNum),
            autoStart:      srLoad("autoStart",      SR_DEFAULTS.autoStart),
            increment:      srLoad("increment",      SR_DEFAULTS.increment),
            handles:        srLoad("handles",        SR_DEFAULTS.handles),
            createNuke:     srLoad("createNuke",     SR_DEFAULTS.createNuke),
            exportXML:      srLoad("exportXML",      SR_DEFAULTS.exportXML),
            createDynLink:  srLoad("createDynLink",  SR_DEFAULTS.createDynLink),
            overscan:       srLoad("overscan",       SR_DEFAULTS.overscan),
            omTemplate:     srLoad("omTemplate",     SR_DEFAULTS.omTemplate),
            shotsFolder:    srLoad("shotsFolder",    SR_DEFAULTS.shotsFolder),
            burninEnabled:  srLoad("burninEnabled",  SR_DEFAULTS.burninEnabled),
            burninProject:  srLoad("burninProject",  SR_DEFAULTS.burninProject),
            burninCompany:  srLoad("burninCompany",  SR_DEFAULTS.burninCompany),
            burninAgency:   srLoad("burninAgency",   SR_DEFAULTS.burninAgency),
            burninClient:   srLoad("burninClient",   SR_DEFAULTS.burninClient)
        });

        // Project-local override: if a Burnin Fields precomp already exists
        // (from a prior roundtrip on this project), its current Source Text
        // values trump the persisted app.settings — matches the "edits in
        // the Burnin Fields comp are sticky" expectation.
        (function overrideFromExistingBurninFields() {
            function readExistingField(fieldName) {
                try {
                    for (var bi = 1; bi <= app.project.numItems; bi++) {
                        var it = app.project.item(bi);
                        if (it instanceof CompItem && it.name === "Burnin Fields") {
                            var L = it.layers.byName(fieldName);
                            if (L && L.text && L.text.sourceText) {
                                return String(L.property("Source Text").value.text || "");
                            }
                        }
                    }
                } catch (e) {}
                return null;
            }
            var eP = readExistingField("Project");
            var eC = readExistingField("Production Company");
            var eA = readExistingField("Agency");
            var eCl = readExistingField("Client");
            if (eP  !== null) etBurninProject.text = eP;
            if (eC  !== null) etBurninCompany.text = eC;
            if (eA  !== null) etBurninAgency.text  = eA;
            if (eCl !== null) etBurninClient.text  = eCl;
        })();

        // ── Buttons ────────────────────────────────────
        var btnGrp = dlg.add("group");
        btnGrp.orientation = "row"; btnGrp.alignment = ["fill", "bottom"]; btnGrp.margins = [0, 4, 0, 0];
        var btnReset   = btnGrp.add("button", undefined, "Reset to Defaults"); btnReset.preferredSize  = [130, 28];
        var btnCancel  = btnGrp.add("button", undefined, "Cancel");            btnCancel.preferredSize = [80,  28];
        var btnSpacer  = btnGrp.add("statictext", undefined, "");              btnSpacer.alignment     = ["fill", "center"];
        var btnOk      = btnGrp.add("button", undefined, "Run Roundtrip");     btnOk.preferredSize     = [140, 28];

        btnReset.onClick  = function() { srApply(SR_DEFAULTS); };
        btnOk.onClick     = function() {
            srSave("prefix",        etPrefix.text);
            srSave("startNum",      etStartNum.text);
            srSave("autoStart",     chkAutoStart.value);
            srSave("increment",     etIncrement.text);
            srSave("handles",       etHandles.text);
            srSave("createNuke",    chkCreateNuke.value);
            srSave("exportXML",     chkExportXML.value);
            srSave("createDynLink", chkCreateDynLink.value);
            // sharedSourceMode is saved from its dedicated preflight dialog
            // on Continue, before Confirm Shots even appears.
            srSave("overscan",      etOverscan.text);
            srSave("omTemplate",    etOM.text);
            srSave("shotsFolder",   etShotsFolder.text);
            srSave("burninEnabled", chkBurninEnabled.value);
            srSave("burninProject", etBurninProject.text);
            srSave("burninCompany", etBurninCompany.text);
            srSave("burninAgency",  etBurninAgency.text);
            srSave("burninClient",  etBurninClient.text);
            dlg.close(1);
        };
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

        // Cancellation bookkeeping — every cancel path feeds reportCancellation()
        // so the user sees WHAT got done before the cancel instead of a bare
        // "cancelled" message. Flags and counters are updated inline as the
        // script makes progress. versionedFile is declared further down; the
        // closure here resolves it at call time so the cancel report can cite
        // the new version filename once it exists.
        var cancelStats = {
            mutationsStarted: false, // flipped to true after saveAsNextVersion succeeds
            shotsCreated:     0,     // shotComps built in the main roundtrip loop
            rendersDone:      0      // plates imported after render
        };
        function reportCancellation(reason) {
            var lines = [reason];
            if (!cancelStats.mutationsStarted) {
                lines.push("", "No changes were made to your project.");
            } else {
                lines.push("");
                if (typeof versionedFile !== "undefined" && versionedFile) {
                    // displayName is URI-decoded (spaces stay spaces);
                    // .name returns percent-encoded ("My%20Project.aep").
                    lines.push("Working file: " + versionedFile.displayName);
                    lines.push("(original preserved on disk as the rollback point)");
                    lines.push("");
                }
                if (cancelStats.shotsCreated > 0) {
                    lines.push("Shot comps built before cancel: " + cancelStats.shotsCreated);
                }
                if (cancelStats.rendersDone > 0) {
                    lines.push("Plates imported before cancel:  " + cancelStats.rendersDone);
                }
                lines.push("", "Use Cmd/Ctrl+Z to undo the session's work, or reopen the original .aep from disk.");
            }
            try { progress.close(); } catch (e) {}

            // ScriptUI window (grey in AE) instead of alert() — matches the
            // Roundtrip Complete summary's look and sidesteps the macOS
            // system-alert styling that the Ae app icon gets slapped onto.
            var w = new Window("dialog", "Roundtrip Cancelled");
            w.orientation = "column"; w.alignChildren = ["fill", "top"];
            w.spacing = 10; w.margins = 14;
            var body = lines.join("\n");
            var bodyLines = body.split("\n").length;
            var maxH = 600;
            try { maxH = $.screens[0].bottom - $.screens[0].top - 200; } catch (eS) {}
            var textH = Math.min(bodyLines * 18 + 10, maxH);
            var txt = w.add("edittext", undefined, body,
                { multiline: true, readonly: true, scrollable: true });
            txt.preferredSize = [480, textH];
            var btnRow = w.add("group"); btnRow.alignment = ["right", "bottom"];
            var ok = btnRow.add("button", undefined, "OK");
            ok.onClick = function () { w.close(); };
            w.show();
        }

        // Cancel-check helper. Call at the top of each long loop iteration:
        //   if (cancelCheck()) return;
        // The return triggers any enclosing try/finally, which correctly ends
        // the active undo group, so the user can Cmd/Ctrl+Z to roll back
        // partial work. Doesn't work during renderQueue.render() — that call
        // blocks the script and AE's own render window takes over, including
        // its own cancel button.
        function cancelCheck() {
            if (!progress.isCancelled()) return false;
            reportCancellation("Roundtrip cancelled by user.");
            return true;
        }
        function pad(n, s) { var str = "" + n; while (str.length < s) str = "0" + str; return str; }

        // Save As → next _v## version before the roundtrip touches anything.
        // VFX convention: MyProject_v03.aep → MyProject_v04.aep. If the current
        // filename has no _v## suffix, tack on _v01. If the target already
        // exists on disk, bump further until we find an unused number. The
        // ORIGINAL file stays on disk untouched as the rollback point; the
        // AE session continues in the new file. Returns the new File on
        // success, or null on failure (caller aborts).
        function saveAsNextVersion() {
            var cur = proj.file;
            if (!cur) return null; // guarded at top of function, defensive
            var baseName = cur.name.replace(/\.aep$/i, "");
            // Underscore before `v` is optional, so both `project_v06` and a
            // bare `v06.aep` are recognised as versioned and bumped in place.
            var m = baseName.match(/^(.*?)(_?v)(\d+)$/);
            var stem, prefix, width, next;
            if (m) {
                stem   = m[1];
                prefix = m[2];
                width  = m[3].length;
                next   = parseInt(m[3], 10) + 1;
            } else {
                stem   = baseName;
                prefix = "_v";
                width  = 2;
                next   = 1;
            }
            var newFile = null;
            while (next < 10000) {
                var candidate = new File(cur.parent.fsName + "/" + stem + prefix + pad(next, width) + ".aep");
                if (!candidate.exists) { newFile = candidate; break; }
                next++;
            }
            if (!newFile) {
                alert("Shot Roundtrip: could not find an unused version number for the backup copy.\nAborting so nothing is modified.");
                return null;
            }
            try {
                // Flush any unsaved edits to the current file first, then
                // Save As to the new version. After proj.save(newFile) the
                // session's current file becomes newFile.
                proj.save();
                proj.save(newFile);
            } catch (eSave) {
                alert("Shot Roundtrip: failed to save versioned copy —\n" + eSave.message +
                      "\n\nAborting so the original file stays untouched.");
                return null;
            }
            return newFile;
        }
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

            // Per-row bake flags. Lives on the entry object so it survives
            // sorts. Toggled via the Toggle Bake button below; default ON
            // so every detected reversed clip bakes unless the user opts
            // out (the bake is the safe choice — keeps a clean forward
            // mainComp + a separate reversed plate for diff-key A/B).
            for (var bf = 0; bf < reversed.length; bf++) {
                if (typeof reversed[bf].bake !== "boolean") reversed[bf].bake = true;
            }

            // Sort controls — ScriptUI listbox headers aren't clickable,
            // so we expose sort as a small button row above the list.
            var revSortKey = 0, revSortDir = 1;
            function revSortList() {
                reversed.sort(function (a, b) {
                    var av, bv;
                    if (revSortKey === 0)      { av = a.layerName;  bv = b.layerName;  }
                    else if (revSortKey === 1) { av = a.path;       bv = b.path;       }
                    else if (revSortKey === 2) { av = a.label;      bv = b.label;      }
                    else                       { av = (a.bake?1:0); bv = (b.bake?1:0); }
                    av = String(av || "").toLowerCase(); bv = String(bv || "").toLowerCase();
                    if (av !== bv) return (av < bv ? -1 : 1) * revSortDir;
                    return 0;
                });
            }
            var revSortRow = w.add("group");
            revSortRow.orientation = "row"; revSortRow.alignChildren = ["left", "center"];
            revSortRow.spacing = 4;
            revSortRow.add("statictext", undefined, "Sort:");
            var REV_LABELS = ["Layer", "Path", "Effect", "Bake"];
            var revBtns = [];
            for (var rs = 0; rs < REV_LABELS.length; rs++) {
                var rb = revSortRow.add("button", undefined, REV_LABELS[rs]);
                rb.preferredSize = [90, 22];
                rb.onClick = (function (col) {
                    return function () {
                        if (revSortKey === col) revSortDir = -revSortDir;
                        else { revSortKey = col; revSortDir = 1; }
                        revSortList();
                        revRepopulate();
                        revRefreshBtns();
                    };
                })(rs);
                revBtns.push(rb);
            }
            function revRefreshBtns() {
                for (var i = 0; i < revBtns.length; i++) {
                    var arrow = (i === revSortKey) ? (revSortDir === 1 ? "  ↓" : "  ↑") : "";
                    revBtns[i].text = REV_LABELS[i] + arrow;
                }
            }

            var lb = w.add("listbox", undefined, [], {
                multiselect: true,
                numberOfColumns: 4, showHeaders: true,
                columnTitles: ["Layer", "Path", "Effect", "Bake"],
                columnWidths: [240, 460, 200, 60]
            });
            function revRepopulate() {
                try { lb.removeAll(); } catch (eRA) {}
                for (var k = 0; k < reversed.length; k++) {
                    var row = lb.add("item", reversed[k].layerName);
                    row.subItems[0].text = reversed[k].path;
                    row.subItems[1].text = reversed[k].label;
                    row.subItems[2].text = reversed[k].bake ? "✓" : "";
                }
            }
            revSortList();
            revRepopulate();
            revRefreshBtns();
            lb.preferredSize = [1000, Math.min(reversed.length * 22 + 40, 420)];

            var warn = w.add("statictext", undefined,
                "Reversed plates break camera tracking, motion vectors, and direction-sensitive VFX \u2014 particle trails, smoke, fire, debris, splashes all look obviously wrong when played backwards. Please verify each one really should ship reversed.",
                { multiline: true });
            warn.preferredSize = [1000, 60];

            var bakeHint = w.add("statictext", undefined,
                "Tip: select one or more rows and click Toggle Bake to render those clips out as new forward-playing plates BEFORE the roundtrip. The layer's source is then swapped to the baked file and time effects are cleared, so re-runs see a clean forward plate.",
                { multiline: true });
            bakeHint.preferredSize = [1000, 50];

            var btnGrp = w.add("group");
            btnGrp.alignment = ["fill", "bottom"];
            btnGrp.margins = [0, 4, 0, 0];
            var btnToggleBake = btnGrp.add("button", undefined, "Toggle Bake");
            btnToggleBake.preferredSize = [120, 28];
            btnToggleBake.helpTip = "Toggle the Bake flag on the selected row(s). (Hotkey: B)";
            var btnBakeAll = btnGrp.add("button", undefined, "Select All");
            btnBakeAll.preferredSize = [100, 28];
            btnBakeAll.helpTip = "Highlight every row in the list. Combine with Toggle Bake to bulk-flip the Bake flag.";
            var btnBakeNone = btnGrp.add("button", undefined, "Deselect All");
            btnBakeNone.preferredSize = [110, 28];
            btnBakeNone.helpTip = "Clear all row selection.";
            var btnSpacer = btnGrp.add("statictext", undefined, "");
            btnSpacer.alignment = ["fill", "center"];
            var btnCancel   = btnGrp.add("button", undefined, "Cancel \u2014 I'll fix first");
            var btnContinue = btnGrp.add("button", undefined, "Continue \u2014 reversed is intentional");
            btnCancel.preferredSize   = [180, 28];
            btnContinue.preferredSize = [260, 28];

            function refreshContinueLabel() {
                var bakeCount = 0;
                for (var rc = 0; rc < reversed.length; rc++) if (reversed[rc].bake) bakeCount++;
                var convCount = reversed.length - bakeCount;
                btnContinue.text = "Continue: " + bakeCount + " bake, " + convCount + " convert";
            }
            refreshContinueLabel();

            function toggleBakeSelection() {
                var sel = lb.selection;
                if (!sel) return;
                var selArr = (sel.length !== undefined) ? sel : [sel];
                var selIndices = [];
                for (var si = 0; si < selArr.length; si++) selIndices.push(selArr[si].index);
                for (var sj = 0; sj < selIndices.length; sj++) {
                    reversed[selIndices[sj]].bake = !reversed[selIndices[sj]].bake;
                    lb.items[selIndices[sj]].subItems[2].text = reversed[selIndices[sj]].bake ? "✓" : "";
                }
                // ScriptUI doesn't repaint subItem text while items are
                // selected, so clear and restore selection.
                lb.selection = null;
                for (var sk = 0; sk < selIndices.length; sk++) lb.items[selIndices[sk]].selected = true;
                refreshContinueLabel();
            }
            function selectAllRows() {
                for (var ai = 0; ai < lb.items.length; ai++) {
                    lb.items[ai].selected = true;
                }
            }
            function deselectAllRows() {
                lb.selection = null;
            }
            btnToggleBake.onClick = toggleBakeSelection;
            btnBakeAll.onClick    = selectAllRows;
            btnBakeNone.onClick   = deselectAllRows;
            try {
                w.addEventListener("keydown", function(e) {
                    if (e.keyName === "B") { e.preventDefault(); toggleBakeSelection(); }
                });
            } catch(eKD) {}

            btnCancel.onClick   = function () { w.close(2); };
            btnContinue.onClick = function () { w.close(1); };

            var dlgRes = w.show();
            if (dlgRes !== 1) return { proceed: false, bakeLayers: [] };
            var bakeLayers = [];
            for (var bi = 0; bi < reversed.length; bi++) {
                if (reversed[bi].bake) bakeLayers.push(reversed[bi]);
            }
            return { proceed: true, bakeLayers: bakeLayers };
        }

        // Shared-source preflight. Fires BEFORE Confirm Shots when the
        // selection contains the same inner footage referenced by
        // multiple selected layers, and lets the user pick a per-clip
        // mode for each shared source.
        //
        // Always returns an object so the caller can persist the
        // dropdown state regardless of cancel/continue (so the user can
        // re-open the dialog later and see their previous picks):
        //   {
        //     cancelled:    true|false,
        //     modes:        { originalSourceId → "separate"|"shared"|"single" }
        //                   (project-session id → mode, used by the runtime),
        //     modesByName:  { sourceName → mode }
        //                   (settings-friendly key → mode, persisted across runs)
        //   }
        // savedByName is the previous run's persisted modesByName, used
        // to initialize each row's dropdown when the source name matches
        // a saved entry — overrides defaultMode for that row.
        //
        // Modes:
        //   separate  this source's references each get their own shot
        //             via chain duplication; one plate per reference.
        //   shared    this source's references collapse into one
        //             shotComp covering the union of every reference's
        //             visible cut; one plate, many edit positions.
        //   single    only the FIRST reference becomes a shot (its chain
        //             is duplicated for isolation); other references
        //             stay untouched and keep pointing at the original
        //             source.
        function confirmSharedSourceMode(groups, defaultMode, savedByName) {
            var w = new Window("dialog", "\u26A0  Shared source footage detected");
            w.orientation = "column"; w.alignChildren = ["fill", "top"];
            w.spacing = 10; w.margins = 14;

            w.add("statictext", undefined,
                groups.length + " source" + (groups.length === 1 ? " is" : "s are") +
                " referenced by multiple selected layers. Pick a mode per source:");

            // Bulk "Set all to" row above the list.
            var bulkGrp = w.add("group");
            bulkGrp.orientation = "row"; bulkGrp.alignChildren = ["left", "center"];
            bulkGrp.spacing = 8;
            bulkGrp.add("statictext", undefined, "Set all to:");
            var btnAllSep = bulkGrp.add("button", undefined, "Separate");
            var btnAllShr = bulkGrp.add("button", undefined, "Shared");
            var btnAllSin = bulkGrp.add("button", undefined, "Single");
            btnAllSep.preferredSize = [90, 22];
            btnAllShr.preferredSize = [90, 22];
            btnAllSin.preferredSize = [90, 22];

            // Per-source rows. ScriptUI's listbox doesn't support inline
            // dropdowns, so build it as a stack of statictext+statictext
            // +dropdown groups instead.
            var listPanel = w.add("panel");
            listPanel.orientation = "column"; listPanel.alignChildren = ["fill", "top"];
            listPanel.spacing = 4; listPanel.margins = [10, 12, 10, 10];

            var headerRow = listPanel.add("group");
            headerRow.orientation = "row"; headerRow.spacing = 8;
            headerRow.alignChildren = ["left", "center"];
            var hSrc = headerRow.add("statictext", undefined, "Source");
            hSrc.preferredSize = [260, 18];
            var hShots = headerRow.add("statictext", undefined, "Shots That Reference It");
            hShots.preferredSize = [380, 18];
            var hMode = headerRow.add("statictext", undefined, "Mode");
            hMode.preferredSize = [120, 18];

            var modeOptions = ["Separate", "Shared", "Single"];
            var defaultIdx = (defaultMode === "shared") ? 1
                           : (defaultMode === "single") ? 2
                           : 0;

            var rowDropdowns = [];
            var savedByNameMap = savedByName || {};
            for (var li = 0; li < groups.length; li++) {
                var lg = groups[li];
                var row = listPanel.add("group");
                row.orientation = "row"; row.spacing = 8;
                row.alignChildren = ["left", "center"];
                var t1 = row.add("statictext", undefined, lg.name);
                t1.preferredSize = [260, 22];
                var t2 = row.add("statictext", undefined, lg.shotNames.join(", "));
                t2.preferredSize = [380, 22];
                var dd = row.add("dropdownlist", undefined, modeOptions);
                dd.preferredSize = [120, 22];
                // Per-source initial selection: previous run's saved pick
                // for THIS source name takes precedence over the global
                // default. Lets the user re-open the dialog after cancel
                // and see exactly what they had.
                var savedForThis = savedByNameMap[lg.name];
                var perRowIdx = (savedForThis === "shared") ? 1
                              : (savedForThis === "single") ? 2
                              : (savedForThis === "separate") ? 0
                              : defaultIdx;
                dd.selection = perRowIdx;
                rowDropdowns.push(dd);
            }

            btnAllSep.onClick = function () { for (var i = 0; i < rowDropdowns.length; i++) rowDropdowns[i].selection = 0; };
            btnAllShr.onClick = function () { for (var i = 0; i < rowDropdowns.length; i++) rowDropdowns[i].selection = 1; };
            btnAllSin.onClick = function () { for (var i = 0; i < rowDropdowns.length; i++) rowDropdowns[i].selection = 2; };

            // Help panel explaining the three modes.
            var helpPanel = w.add("panel", undefined, "What these modes do");
            helpPanel.orientation = "column"; helpPanel.alignChildren = ["fill", "top"];
            helpPanel.spacing = 4; helpPanel.margins = [10, 12, 10, 10];

            function helpRow(label, desc) {
                var hr = helpPanel.add("group");
                hr.orientation = "row"; hr.alignChildren = ["left", "top"]; hr.spacing = 8;
                var hl = hr.add("statictext", undefined, label);
                hl.preferredSize = [80, 38];
                try {
                    hl.graphics.font = ScriptUI.newFont(hl.graphics.font.name, "BOLD", hl.graphics.font.size);
                } catch (eHF) {}
                var hd = hr.add("statictext", undefined, desc, { multiline: true });
                hd.preferredSize = [700, 38];
            }
            helpRow("Separate",
                "Each reference becomes its own shot. The precomp chain is duplicated per reference so each gets its own plate, " +
                "shotComp, and disk folder. Pick this when you want to grade / VFX the same source differently per occurrence in the edit.");
            helpRow("Shared",
                "All references collapse into ONE shotComp. Its render covers the union of every reference's visible cut, so no frames are " +
                "truncated. Every reference plays the same plate at its own edit position. Pick this when you want one consistent treatment everywhere.");
            helpRow("Single",
                "Only the FIRST reference becomes a shot (its chain is isolated via duplication). Other references in mainComp stay untouched " +
                "and keep pointing at the original source. Pick this when you only want to roundtrip ONE occurrence and leave the others as-is.");

            var btnGrp = w.add("group");
            btnGrp.alignment = ["right", "bottom"];
            btnGrp.margins = [0, 4, 0, 0];
            var btnCancel   = btnGrp.add("button", undefined, "Cancel");
            var btnContinue = btnGrp.add("button", undefined, "Continue");
            btnCancel.preferredSize   = [100, 28];
            btnContinue.preferredSize = [120, 28];
            btnCancel.onClick   = function () { w.close(2); };
            btnContinue.onClick = function () { w.close(1); };

            var dlgRet = w.show();
            // Capture selections regardless of cancel/continue so the
            // caller can persist the dropdown state for the next run.
            var modes = {};
            var modesByName = {};
            for (var di = 0; di < groups.length; di++) {
                var sel = rowDropdowns[di].selection ? rowDropdowns[di].selection.index : 0;
                var modeStr = (sel === 1) ? "shared" : (sel === 2) ? "single" : "separate";
                modes[groups[di].originalSourceId] = modeStr;
                modesByName[groups[di].name] = modeStr;
            }
            return {
                cancelled:   (dlgRet !== 1),
                modes:       modes,
                modesByName: modesByName
            };
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

                    // Capture the original layer's label colour so we can
                    // restore it on the new wrapper layer in mainComp.
                    // AE's precompose otherwise hands the wrapper layer
                    // a default label that doesn't match anything the
                    // user might have set up (e.g. the "Color Time-
                    // Reverse Layers" pass colours reversed clips blue).
                    var aOriginalLabel = 0;
                    try { aOriginalLabel = comp.layer(aInfo.index).label; } catch (eLbl0) {}

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
                        try { aPrecompLayer.label = aOriginalLabel; } catch (eLbl1) {}
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

        // Recursive: does this comp (or any nested precomp under it) have
        // a layer whose source is `srcItem`? Used by the bake-placement
        // step to find which shot owns a given baked source. `visited` is
        // a {compName: true} cycle-guard.
        function compTreeContainsSource(comp, srcItem, visited) {
            if (!comp || !(comp instanceof CompItem)) return false;
            visited = visited || {};
            var key = "c" + (comp.id || comp.name);
            if (visited[key]) return false;
            visited[key] = true;
            for (var li = 1; li <= comp.numLayers; li++) {
                var L;
                try { L = comp.layer(li); } catch (eL1) { continue; }
                if (!L) continue;
                try {
                    if (L.source === srcItem) return true;
                    if (L.source instanceof CompItem) {
                        if (compTreeContainsSource(L.source, srcItem, visited)) return true;
                    }
                } catch (eL2) {}
            }
            return false;
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
        // currently visible. Returns { footageLayer, footageComp, breadcrumb, layerChain }.
        //
        // layerChain holds the precomp LAYERS encountered between the caller's starting
        // comp and the footage's immediate parent, outer-to-inner. Used by the expansion
        // step to walk the outer selection's in/out through every nested stretch + time
        // remap via mapTimeToSource — so the render range reflects what's actually visible
        // in mainComp, not the footage layer's untrimmed extent.
        //
        //   mainComp(selected precomp A)
        //     └─ precomp A
        //         └─ B_layer → precomp B        ← layerChain = [B_layer]
        //             └─ footage.mov            ← footageLayer, footageComp = B
        //
        // For footage directly inside the selected precomp, layerChain is empty.
        function findAllFootageInPrecomp(comp, path, layerChain) {
            var currentPath  = (path       || []).concat([comp.name]);
            var currentChain =  layerChain || [];
            var results = [];
            for (var li = 1; li <= comp.numLayers; li++) {
                var l = comp.layer(li);
                if (!l.hasVideo || l.guideLayer || l.adjustmentLayer || l.nullLayer) continue;
                if (l.source === null) continue;
                // Footage file?
                var isFile = false;
                try { if (l.source.mainSource && l.source.mainSource.file) isFile = true; } catch(e) {}
                if (isFile) {
                    results.push({
                        footageLayer: l,
                        footageComp:  comp,
                        breadcrumb:   currentPath,
                        layerChain:   currentChain.slice()  // snapshot so siblings don't share
                    });
                } else if (l.source instanceof CompItem) {
                    // Sub-precomp: recurse and extend the chain with THIS precomp layer,
                    // since future time mappings have to travel through it.
                    var sub = findAllFootageInPrecomp(l.source, currentPath, currentChain.concat([l]));
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
        var NUKE_PROFILE   = "ProRes 4:2:2 HQ 10-bit";
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
        // shotFPS is the source PLATE's native frame rate (may differ from the
        // mainComp / AppendClip delivery rate). Using plate-fps at Root.fps
        // keeps Read/Write frame counts and durations in sync — using the
        // mainComp fps here would retime every 50fps plate as 25fps on output.
        function writeNukeShotScript(shotFile, d) {
            var shotFPS   = d.fps || 25;
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
            nk += "# Format:    " + d.w + "x" + d.h + " @ " + shotFPS + " fps\n";
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
            nk += "#    \"" + d.name + "_comp_v03.mov\"). The AE \"Import Returns\"\n";
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
            nk += " fps " + shotFPS + "\n";
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
        var autoStartMode = !!chkAutoStart.value;
        var handleFrames = parseInt(etHandles.text, 10);
        var overscanPercent = parseFloat(etOverscan.text); if(isNaN(overscanPercent)) overscanPercent=0;
        var omTemplate = etOM.text;
        var increment = parseInt(etIncrement.text, 10);
        if (isNaN(increment) || increment < 1) increment = 10;

        // Auto-numbering: scan mainComp for existing shot layers (named
        // {prefix}NNN_comp, {prefix}NNN_container, or range-bin
        // {prefix}NNN_MMM_container) and assemble their (number, time)
        // positions + a taken-numbers map. Used by pickAutoShotNumber()
        // below to sandwich each new shot between its time-neighbours.
        var autoExisting = [];  // [{num, time}, …] sorted by time
        var autoTaken    = {};  // {num: true} for collision avoidance
        if (autoStartMode) {
            var aPrefEsc = shotPrefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            // Project-wide: any {prefix}NNN_comp or {prefix}NNN_container marks NNN as taken.
            var aCompRe      = new RegExp("^" + aPrefEsc + "(\\d+)_(?:comp|container)(?:_OS)?$", "i");
            var aRangeRe     = new RegExp("^" + aPrefEsc + "(\\d+)_(\\d+)_container(?:_OS)?$", "i");
            for (var aPi = 1; aPi <= proj.numItems; aPi++) {
                var aIt = proj.item(aPi);
                if (!(aIt instanceof CompItem)) continue;
                var aM = aIt.name.match(aCompRe);
                if (aM) { autoTaken[parseInt(aM[1], 10)] = true; continue; }
                var aR = aIt.name.match(aRangeRe);
                if (aR) {
                    // Range container spans [first, last] — mark every integer
                    // in the range as taken so sandwich picks can't collide
                    // with a middle number (e.g. shot_010_030_container
                    // implies 020 also exists).
                    var aRF = parseInt(aR[1], 10);
                    var aRL = parseInt(aR[2], 10);
                    for (var aRx = Math.min(aRF, aRL); aRx <= Math.max(aRF, aRL); aRx++) {
                        autoTaken[aRx] = true;
                    }
                }
            }
            // mainComp layers: record (num, inPoint) for time-order sandwich detection.
            var aTimeRe = new RegExp("^" + aPrefEsc + "(\\d+)(?:_(\\d+))?_(?:comp|container)(?:_OS)?$", "i");
            for (var aLi = 1; aLi <= mainComp.numLayers; aLi++) {
                var aL = mainComp.layer(aLi);
                var aNm = null;
                try { if (aL.source instanceof CompItem) aNm = aL.source.name; } catch(eAN) {}
                if (!aNm) continue;
                var aTM = aNm.match(aTimeRe);
                if (!aTM) continue;
                var aT = 0;
                try { aT = aL.inPoint; } catch(eAT) {}
                autoExisting.push({ num: parseInt(aTM[1], 10), time: aT });
                if (aTM[2]) autoExisting.push({ num: parseInt(aTM[2], 10), time: aT });
            }
            // Stable order at tied times: secondary key = num ascending.
            // ExtendScript's Array.sort is NOT guaranteed stable, so ties
            // (like range-container's first+last entries sharing one
            // layer's inPoint) can reorder unpredictably. The time-ordered
            // scan picks the LAST entry with time <= t as `before`, so a
            // reorder between {num:10} and {num:30} at the same time would
            // flip `before` from 30 to 10 and produce a wrong midpoint.
            autoExisting.sort(function(a, b) {
                if (a.time !== b.time) return a.time - b.time;
                return a.num - b.num;
            });
        }

        // Pick a shot number for a layer at mainComp time t. Neighbour-
        // aware: sandwiches between the existing before/after shot numbers,
        // so inserting a new layer between shot_030 and shot_040 gives
        // shot_035. Collisions (midpoint already taken) are resolved by
        // bumping upward to the next free number.
        function pickAutoShotNumber(t) {
            var before = null, after = null;
            for (var i = 0; i < autoExisting.length; i++) {
                if (autoExisting[i].time <= t) before = autoExisting[i];
                else { after = autoExisting[i]; break; }
            }
            var pick;
            if (!before && !after)      pick = increment;
            else if (!before)           pick = Math.max(1, after.num - increment);
            else if (!after)            pick = before.num + increment;
            else                        pick = Math.floor((before.num + after.num) / 2);
            while (autoTaken[pick] && pick < 100000) pick++;
            autoTaken[pick] = true;
            return pick;
        }

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

        // Open the progress palette now — the preflight scan and Confirm
        // Shots prep use it for feedback. The Save-As _v## bump is deferred
        // until AFTER Confirm Shots is accepted (see below), so that
        // cancelling on any preflight dialog leaves the user's original
        // .aep untouched on disk — no empty-change version files left over.
        var versionedFile = null;
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
                if (fx) results.push({ layerName: l.name, path: path.join(" \u203A "), label: fx.label, reversed: fx.reversed, topLevel: false, layer: l });
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
            if (topFx) results.push({ layerName: tl.name, path: mainComp.name, label: topFx.label, reversed: topFx.reversed, topLevel: true, layer: tl });
            if (tl.source instanceof CompItem) {
                walkPrecompForEffects(tl.source, [mainComp.name, tl.source.name], results);
            }
            return results;
        }

        // Tolerance for matching auto-keys vs our deliberately-placed keys
        // during the prune step. About 1/2000 of a second — well below any
        // realistic frame rate (240fps has 1/240 ≈ 4.2ms spacing) but loose
        // enough to catch AE's own floating-point drift on key insertion.
        var KEY_TIME_EPSILON_SEC = 0.0005;

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
        //
        // SYNC: helpers/reverse_stretch_to_remap.jsx has an identical copy of
        // this algorithm. Keep them in lockstep on bug fixes.
        function convertStretchReversalToRemap(layer) {
            try {
                if (layer.timeRemapEnabled) return false;
                if (!(layer.stretch < 0)) return false;
                var srcDur = (layer.source && layer.source.duration) ? layer.source.duration : 0;
                if (srcDur <= 0) return false;

                // For negatively-stretched layers AE swaps the reported in/out:
                //   inPoint  = comp-LATER  edge (source-earlier side)
                //   outPoint = comp-EARLIER edge (source-later side)
                // Normalize into comp-timeline order before sampling source times.
                var startT   = layer.startTime;
                var stretch  = layer.stretch;
                var rawIn    = layer.inPoint;
                var rawOut   = layer.outPoint;
                var frameDur = layer.containingComp.frameDuration;

                var compStart, compEnd;
                if (rawIn <= rawOut) { compStart = rawIn;  compEnd = rawOut; }
                else                 { compStart = rawOut; compEnd = rawIn;  }

                // sourceTime(compTime) = (compTime - startTime) * (100 / stretch).
                // For reversed layers AE anchors startTime one frame past the last
                // rendered source frame, so subtract one frameDur to land on the
                // source time that's actually on screen.
                //
                // Sample at the FIRST (compStart) and LAST (compEnd - frameDur)
                // rendered comp frames — not the layer edges. outPoint is
                // exclusive, so compEnd itself is never rendered; pairing srcAtEnd
                // at compEnd with the (cutDur - frameDur) slope denominator below
                // would mismatch by one frameDur and drift the end of playback.
                var srcAtStart = (compStart            - startT) * (100 / stretch) - frameDur;
                var srcAtEnd   = (compEnd  - frameDur  - startT) * (100 / stretch) - frameDur;

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

                // Four keys all on the same line (the actual playback curve).
                // Cut-boundary keys pin the cut range at the intended speed;
                // the outer two extend into the handle range so AE doesn't
                // extrapolate past the last key in *_dynamicLink wrappers or
                // anywhere else the layer gets extended.
                //
                // Slope = (srcAtEnd - srcAtStart) / (cutDur - frameDur) — the
                // actual stretch rate between the first and last rendered comp
                // frames. Captures any negative-stretch magnitude (-100 → -1,
                // -50 → -2, -200 → -0.5, -124 → ≈ -0.806).
                //
                // endKeyTime sits AT compEnd (on the cut_out marker) rather
                // than at the last rendered frame. endKeyVal is extrapolated
                // one frameDur past srcAtEnd along the slope, so the key sits
                // on the marker while LINEAR interpolation between (compStart,
                // srcAtStart) and (compEnd, endKeyVal) still hits srcAtEnd
                // exactly at the last rendered frame (compEnd - frameDur).
                // Don't lower-clamp endKeyVal: extrapolating past source edge
                // is expected here (for full-clip reversal endKeyVal lands at
                // about -frameDur) and clamping it to 0 would break the slope.
                var endKeyTime = compEnd;
                var cutDurLen  = compEnd - compStart;
                var handleSec  = (layer.containingComp && layer.containingComp.frameRate)
                               ? (handleFrames / layer.containingComp.frameRate)
                               : 0;
                var slope      = ((cutDurLen - frameDur) > 0)
                               ? ((srcAtEnd - srcAtStart) / (cutDurLen - frameDur))
                               : 0;
                var preTime    = compStart - handleSec;
                var postTime   = compEnd   + handleSec;
                var preVal     = srcAtStart - slope * handleSec;
                var endKeyVal  = srcAtEnd   + slope * frameDur;
                var postVal    = srcAtEnd   + slope * (frameDur + handleSec);
                if (preVal    < 0)      preVal    = 0;
                if (preVal    > srcDur) preVal    = srcDur;
                if (endKeyVal > srcDur) endKeyVal = srcDur;
                if (postVal   > srcDur) postVal   = srcDur;

                // Write our keys first, then prune AE's auto-keys — clearing
                // them before writing can leave the property in an unusable state.
                tr.setValueAtTime(preTime,    preVal);
                tr.setValueAtTime(compStart,  srcAtStart);
                tr.setValueAtTime(endKeyTime, endKeyVal);
                tr.setValueAtTime(postTime,   postVal);

                for (var k = tr.numKeys; k >= 1; k--) {
                    var kt = tr.keyTime(k);
                    if (Math.abs(kt - preTime)    > KEY_TIME_EPSILON_SEC &&
                        Math.abs(kt - compStart)  > KEY_TIME_EPSILON_SEC &&
                        Math.abs(kt - endKeyTime) > KEY_TIME_EPSILON_SEC &&
                        Math.abs(kt - postTime)   > KEY_TIME_EPSILON_SEC) {
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


        // Scan for reversals FIRST so the warning dialog comes up BEFORE any
        // DOM mutation. Cancel on this dialog must leave the project
        // unmodified — earlier versions converted negative-stretch to
        // time-remap before the dialog, so a cancel left the conversions
        // already applied.
        progress.update("Scanning for reversed clips\u2026", "", 3);
        var reversedList = []; // drives the loud confirm dialog
        for (var ri = 0; ri < selLayers.length; ri++) {
            if (cancelCheck()) return;
            var rTl = selLayers[ri].layer;
            var rEffects = scanSelLayerForEffects(rTl);
            for (var rE = 0; rE < rEffects.length; rE++) {
                if (rEffects[rE].reversed) reversedList.push(rEffects[rE]);
            }
        }

        // Loud dialog only when a reversal is present. Non-reversed time
        // effects proceed silently and show up as passive info next to the
        // shot in the Confirm Shots dialog.
        //
        // IMPORTANT: this dialog is ADVISORY — read-only. The actual
        // stretch→remap conversion and auto-precompose happen much later,
        // after the user commits on the Confirm Shots preflight. Hitting
        // "Continue" here only acknowledges the warning; cancelling on
        // EITHER this dialog OR Confirm Shots leaves the project pristine.
        //
        // Layers picked for "render as new plate" in the warning dialog
        // are stashed on `revBakeLayers` and consumed by the post-Save-As
        // bake phase below.
        var revBakeLayers = [];
        if (reversedList.length > 0) {
            // Close palette before modal so macOS hands focus to the dialog.
            progress.close();
            var revRes = confirmReversedClips(reversedList);
            progress = makeProgressPanel();
            if (!revRes || !revRes.proceed) {
                reportCancellation("Cancelled at the reversed-clips warning \u2014 no roundtrip performed.");
                return;
            }
            revBakeLayers = revRes.bakeLayers || [];
            progress.update("Reversed clips acknowledged, preparing preflight\u2026", "", 4);
        }

        // Shared expansion: called once pre-preflight to populate the
        // Confirm Shots dialog from the ORIGINAL (unconverted, un-wrapped)
        // layer state, and again post-accept after conversion + auto-
        // precompose to rebuild against the new container layer refs.
        // Returns null if the user canceled mid-walk.
        function buildExpandedLayers(selLayersIn, skippedLayersIn, progBaseStart, progBaseSpan, showProgress) {
            var out = [];
            for (var ei = 0; ei < selLayersIn.length; ei++) {
                if (cancelCheck()) return null;
                if (showProgress) {
                    progress.update(null,
                        (ei + 1) + " of " + selLayersIn.length + " selected layers walked",
                        progBaseStart + (progBaseSpan * ei / Math.max(1, selLayersIn.length)));
                }
                var eItem = selLayersIn[ei];
                if (!eItem.isPrecomp) {
                    out.push({
                        layer: eItem.layer, mainLayerIdx: eItem.layer.index, isPrecomp: false, found: null, totalInPrecomp: 0,
                        // Snapshot the source id NOW, before any replaceSource runs later in the
                        // processing loop — dedup must key on the original source, not whatever
                        // shotComp has taken its place.
                        originalSourceId: (eItem.layer.source && eItem.layer.source.id) ? eItem.layer.source.id : null
                    });
                } else {
                    var eFounds = findAllFootageInPrecomp(eItem.layer.source);
                    if (eFounds.length === 0) {
                        skippedLayersIn.push(eItem.layer.name + " (no footage found inside precomp)");
                    } else {
                        for (var ef = 0; ef < eFounds.length; ef++) {
                            var eF = eFounds[ef];
                            // Per-footage source range = intersection of two signals:
                            //
                            //   1. CHAIN-DERIVED range: walk the outer's visible in/out
                            //      down through every precomp layer + the footage layer
                            //      via mapTimeToSource. Produces the source time span
                            //      the edit COULD show for this footage — but because
                            //      mapTimeToSource only applies stretch + time-remap
                            //      (not inPoint/outPoint trim), it extrapolates past
                            //      the footage layer's own cut. In a multi-footage
                            //      precomp this ends up as the full outer span (e.g.
                            //      C2454 trimmed 0-10 in a 0-110 precomp still returns
                            //      0-110 source).
                            //
                            //   2. OWN range: getRequiredSourceRange reads the footage
                            //      layer's own inPoint/outPoint — the cut region it
                            //      actually occupies inside its parent precomp.
                            //
                            // Intersecting gives the minimal source range that is both
                            // VISIBLE from mainComp AND WITHIN this layer's own cut.
                            // Covers the user's original ask (2-sec cut of 2-min clip
                            // buried N precomps deep → render 2 seconds) without
                            // breaking multi-footage precomps (each footage layer only
                            // renders its own trimmed cut).
                            var rA = mapTimeToSource(eItem.layer, eItem.layer.inPoint);
                            var rB = mapTimeToSource(eItem.layer, eItem.layer.outPoint);
                            for (var lcI = 0; lcI < eF.layerChain.length; lcI++) {
                                rA = mapTimeToSource(eF.layerChain[lcI], rA);
                                rB = mapTimeToSource(eF.layerChain[lcI], rB);
                            }
                            rA = mapTimeToSource(eF.footageLayer, rA);
                            rB = mapTimeToSource(eF.footageLayer, rB);
                            var outerMin = Math.min(rA, rB);
                            var outerMax = Math.max(rA, rB);
                            var ownRange = getRequiredSourceRange(eF.footageLayer);
                            var eRange = {
                                start: Math.max(ownRange.start, outerMin),
                                end:   Math.min(ownRange.end,   outerMax)
                            };
                            // Fallback if intersection is empty (footage isn't truly
                            // visible from the outer range, but findAllFootageInPrecomp
                            // still surfaced it — e.g. it sits entirely before/after
                            // the outer's trim window). Render its own range instead
                            // of creating a zero-duration shot.
                            if (eRange.start >= eRange.end) {
                                eRange.start = ownRange.start;
                                eRange.end   = ownRange.end;
                            }
                            out.push({
                                layer: eItem.layer, mainLayerIdx: eItem.layer.index, isPrecomp: true, totalInPrecomp: eFounds.length,
                                found: { footageLayer: eF.footageLayer, footageComp: eF.footageComp,
                                         rangeStart: eRange.start, rangeEnd: eRange.end,
                                         breadcrumb: eF.breadcrumb || [],
                                         layerChain: eF.layerChain || [] },
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
            return out;
        }

        // Deep-duplicate the precomp chain for an expandedLayers entry so
        // subsequent replaceSource mutations on its inner footage layer
        // don't leak back to any shared precomps in the original chain.
        //
        // Used by the cross-selection isolation pre-pass below: when the
        // same footage source is referenced by two different selected
        // layers (e.g. Cloud Space slow placed in mainComp twice), the
        // SECOND entry gets its entire chain duplicated so it can become
        // its own separate shot without stomping on the first entry's
        // replaceSource. AE's CompItem.duplicate() is shallow — the
        // duplicated comp's INNER layers are new Layer objects but their
        // `.source` references still point at the same sub-precomps. For
        // deeper chains we recurse: at each chain level, duplicate the
        // sub-precomp too and replaceSource the layer so each level is
        // isolated end-to-end.
        //
        // Mutates: item.layer.source (replaceSource to outer duplicate),
        // item.found.footageLayer / .footageComp / .layerChain (rebound
        // into the deepest duplicate).
        function duplicateChainForItem(item) {
            if (!item || !item.isPrecomp || !item.found) return;
            var outerSource;
            try { outerSource = item.layer.source; } catch(eOS) { return; }
            if (!outerSource) return;
            var outerDup;
            try { outerDup = outerSource.duplicate(); } catch(eDup) { return; }
            try { item.layer.replaceSource(outerDup, false); } catch(eRS) { return; }

            var chain = item.found.layerChain || [];
            // File the duplicates under /Shots/ so they don't clutter the
            // project root where the user may have organised their comps.
            try {
                var dupBin = getShotBin(getBinFolder("Shots"), "autoPrecomps");
                outerDup.parentFolder = dupBin;
            } catch(eBinOD) {}

            if (chain.length === 0) {
                // Footage lives directly inside the outer. Rebind into the
                // duplicate's equivalent layer (layer indices are preserved
                // across CompItem.duplicate()).
                var fIdx0;
                try { fIdx0 = item.found.footageLayer.index; } catch(eFI0) { return; }
                item.found.footageComp = outerDup;
                try { item.found.footageLayer = outerDup.layer(fIdx0); } catch(eFL0) {}
                return;
            }

            // Walk the chain, duplicating each sub-precomp and rebinding
            // the path so the final footage layer lives in a fully-owned
            // clone tree.
            var currentComp = outerDup;
            var newChain = [];
            for (var ci = 0; ci < chain.length; ci++) {
                var origIdx;
                try { origIdx = chain[ci].index; } catch(eCI) { return; }
                var layerInCurrent;
                try { layerInCurrent = currentComp.layer(origIdx); } catch(eLI) { return; }
                var subDup;
                try { subDup = layerInCurrent.source.duplicate(); } catch(eSD) { return; }
                try { layerInCurrent.replaceSource(subDup, false); } catch(eSRS) { return; }
                try { subDup.parentFolder = dupBin; } catch(eBSD) {}
                newChain.push(layerInCurrent);
                currentComp = subDup;
            }

            item.found.footageComp = currentComp;
            item.found.layerChain  = newChain;
            var fIdxN;
            try { fIdxN = item.found.footageLayer.index; } catch(eFIN) { return; }
            try { item.found.footageLayer = currentComp.layer(fIdxN); } catch(eFLN) {}
        }

        // Shared finalize step: conversion + auto-precompose + re-scan.
        // Mutates the DOM — only called AFTER the user commits on the
        // Confirm Shots preflight dialog. Returns the new selLayers, or
        // null on abort/failure (caller returns early).
        function applyTimeEffectConversions() {
            // Snapshot the mainComp selection — the converter flips layer/
            // property selection flags to dodge the "hidden property" error,
            // which collapses the user's original selection in the UI.
            // Restore it after the loop.
            var preConvertSelection = [];
            for (var pcs = 0; pcs < mainComp.selectedLayers.length; pcs++) {
                preConvertSelection.push(mainComp.selectedLayers[pcs]);
            }
            progress.update("Converting any reversed clips to time remaps\u2026", "0 of " + selLayers.length, 16);
            for (var cri = 0; cri < selLayers.length; cri++) {
                if (cancelCheck()) return null;
                progress.update(null,
                    "layer " + (cri + 1) + " of " + selLayers.length,
                    16 + 2 * (cri / Math.max(1, selLayers.length)));
                var crl = selLayers[cri].layer;
                if (!isScanRelevantLayer(crl)) continue;
                convertStretchReversalToRemap(crl);
                if (crl.source instanceof CompItem) {
                    convertAllStretchReversalsInComp(crl.source);
                }
            }
            try {
                for (var dse = 1; dse <= mainComp.numLayers; dse++) mainComp.layer(dse).selected = false;
                for (var rse = 0; rse < preConvertSelection.length; rse++) {
                    try { preConvertSelection[rse].selected = true; } catch(eSelRestore) {}
                }
            } catch(eSelSnap) {}

            // Build trAffected AFTER the convert so inPoint/outPoint reflect
            // the post-conversion forward-ordered values
            // (convertStretchReversalToRemap swaps them into comp-timeline
            // order when re-anchoring a reversed layer). autoPrecomposeTrimmed
            // expects forward in/out.
            var trAffected = [];
            for (var ti = 0; ti < selLayers.length; ti++) {
                var tl = selLayers[ti].layer;
                var topFx = describeTimeEffect(tl);
                if (topFx) {
                    trAffected.push({ selIdx: ti, index: tl.index, name: tl.name, inPoint: tl.inPoint, outPoint: tl.outPoint,
                        label: topFx.label, reversed: topFx.reversed });
                }
            }

            if (trAffected.length > 0) {
                progress.update("Auto-precomposing " + trAffected.length + " time-remapped layer(s)\u2026", "", 18);
                if (!autoPrecomposeTrimmed(mainComp, trAffected)) {
                    // autoPrecomposeTrimmed returns false on internal error OR
                    // on cancel. Surface the cancel message if that's why we
                    // stopped.
                    if (cancelCheck()) return null;
                    return null;
                }

                // Re-scan the selection — the precomposed layers replaced the
                // originals. Apply the same filter as the initial scan:
                // auto-precompose's selection restoration puts every pre-
                // existing selection back on the layer panel, including
                // shape/text/null/adjustment/guide layers the user happened
                // to have selected alongside the real shots. Without this
                // filter those come through as roundtrip candidates and
                // crash downstream on null .source.
                var newSel = [];
                for (var ri = 1; ri <= mainComp.numLayers; ri++) {
                    if (!mainComp.layer(ri).selected) continue;
                    var rLayer = mainComp.layer(ri);
                    if (!rLayer.hasVideo || rLayer.guideLayer || rLayer.adjustmentLayer || rLayer.nullLayer) continue;
                    if (rLayer.source === null) { skippedLayers.push(rLayer.name + " (Shape/Text)"); continue; }
                    var rIsFile = false;
                    try { if (rLayer.source.mainSource && rLayer.source.mainSource.file) rIsFile = true; } catch(eRIS) {}
                    var rIsPrecomp = (rLayer.source instanceof CompItem);
                    if (rIsFile || rIsPrecomp) {
                        newSel.push({ layer: rLayer, isPrecomp: rIsPrecomp, mainLayerIdx: rLayer.index });
                    } else {
                        skippedLayers.push(rLayer.name + " (Solid/Shape/Text)");
                    }
                }
                if (newSel.length === 0) {
                    alert("No layers selected after auto-precompose. Please select the precomposed layers and try again.");
                    return null;
                }
                return newSel;
            }
            return selLayers;
        }

        // Expand selected layers into shot entries (ORIGINAL, unconverted
        // state — the conversion + auto-precompose runs only after the
        // Confirm Shots dialog is accepted, so this preview reflects
        // exactly what the user currently has in their comp).
        progress.update("Expanding precomps and resolving source ranges\u2026", "0 of " + selLayers.length, 11);
        var expandedLayers = buildExpandedLayers(selLayers, skippedLayers, 11, 5, true);
        if (expandedLayers === null) return;

        // If auto-numbering is on, assign a number per expandedLayers entry
        // ONCE up front (shared between preview and main loop). Multi-footage
        // precomps share a mainLayerIdx, so the first entry for that idx gets
        // the time-based pick and sub-entries step forward by `increment`.
        var autoShotNumbers = null;
        if (autoStartMode) {
            autoShotNumbers = [];
            var autoBaseByMainIdx  = {};
            var autoCountByMainIdx = {};
            function autoRegisterAssigned(num, time) {
                // Make every newly-assigned shot visible to later picks as a
                // time-neighbour — otherwise selections after the first all
                // see autoExisting=[] and pick `increment`, then cascade
                // upward through collision-bumps (10, 11, 12, …) instead of
                // spacing out by increment from their time-nearest predecessor.
                autoExisting.push({ num: num, time: time });
                autoExisting.sort(function(a, b) {
                    if (a.time !== b.time) return a.time - b.time;
                    return a.num - b.num;
                });
            }
            for (var aEi = 0; aEi < expandedLayers.length; aEi++) {
                var aEItem = expandedLayers[aEi];
                var aMLidx = aEItem.mainLayerIdx;
                var aLT = 0;
                try { aLT = aEItem.layer.inPoint; } catch(eALT) {}
                var aAssigned;
                if (!(aMLidx in autoBaseByMainIdx)) {
                    autoBaseByMainIdx[aMLidx]  = pickAutoShotNumber(aLT);
                    autoCountByMainIdx[aMLidx] = 0;
                    aAssigned = autoBaseByMainIdx[aMLidx];
                } else {
                    autoCountByMainIdx[aMLidx]++;
                    aAssigned = autoBaseByMainIdx[aMLidx] + autoCountByMainIdx[aMLidx] * increment;
                    while (autoTaken[aAssigned] && aAssigned < 100000) aAssigned++;
                    autoTaken[aAssigned] = true;
                }
                autoRegisterAssigned(aAssigned, aLT);
                autoShotNumbers.push(aAssigned);
            }
        }

        if (expandedLayers.length === 0) {
            var msg = "Keine gültigen Layer gefunden.";
            if (skippedLayers.length > 0) msg += "\n" + skippedLayers.length + " Layer wurden übersprungen.";
            alert(msg); progress.close(); return;
        }

        // Cross-selection shared-source detection. An expandedLayers entry
        // is "shared" when another entry shares its originalSourceId BUT
        // comes from a different mainLayerIdx (selected top-level layer).
        // Intra-selection repeats (same outer precomp cutting one source
        // multiple times) aren't flagged — those are expected to expand
        // into N shots.
        var sharedByOrig = {}; // originalSourceId → { name, shotNames[], mainLayerIdxs{} }
        for (var sgi = 0; sgi < expandedLayers.length; sgi++) {
            var sgItem = expandedLayers[sgi];
            if (!sgItem.originalSourceId) continue;
            var sgNum0  = autoShotNumbers ? autoShotNumbers[sgi] : (startNum + sgi * increment);
            var sgShot0 = shotPrefix + pad(sgNum0, 3);
            var sgSrc0;
            try {
                sgSrc0 = sgItem.isPrecomp
                    ? (sgItem.found.footageLayer.source && sgItem.found.footageLayer.source.name)
                    : (sgItem.layer.source && sgItem.layer.source.name);
            } catch(eSS) { sgSrc0 = null; }
            var grp0 = sharedByOrig[sgItem.originalSourceId];
            if (!grp0) {
                grp0 = {
                    originalSourceId: sgItem.originalSourceId,
                    name:             sgSrc0 || "(no source)",
                    shotNames:        [],
                    mainLayerIdxs:    {}
                };
                sharedByOrig[sgItem.originalSourceId] = grp0;
            }
            grp0.shotNames.push(sgShot0);
            grp0.mainLayerIdxs[sgItem.mainLayerIdx] = true;
        }
        var sharedGroups = {};
        var sharedGroupsList = [];
        for (var sgK0 in sharedByOrig) {
            if (!sharedByOrig.hasOwnProperty(sgK0)) continue;
            var sgG0 = sharedByOrig[sgK0];
            var sgDistinct = 0;
            for (var sgK1 in sgG0.mainLayerIdxs) if (sgG0.mainLayerIdxs.hasOwnProperty(sgK1)) sgDistinct++;
            if (sgDistinct >= 2) {
                sharedGroups[sgK0] = sgG0;
                sharedGroupsList.push(sgG0);
            }
        }

        // Shared-source preflight dialog. Fires BEFORE Confirm Shots when
        // sharing is detected, mirroring the reversed-clips warning's
        // shape. Captures a per-clip mode (separate / shared / single)
        // for each shared source. The post-accept pre-pass looks up the
        // mode per group when deciding what to do.
        var sharedSourceModes = {}; // originalSourceId → "separate"|"shared"|"single"
        if (sharedGroupsList.length > 0) {
            var sharedSourceModeDefault = srLoad("sharedSourceMode", SR_DEFAULTS.sharedSourceMode);
            if (sharedSourceModeDefault !== "shared" && sharedSourceModeDefault !== "single") {
                sharedSourceModeDefault = "separate";
            }
            // Per-source picks from previous runs, keyed by source name.
            // Used to pre-select dropdowns so the user sees their last
            // choices when re-opening the dialog (whether they cancelled
            // or continued before).
            var savedByNamePrev = {};
            try {
                var rawByName = srLoad("sharedSourceModesByName", "");
                if (rawByName) {
                    var parsed = JSON.parse(rawByName);
                    if (parsed && typeof parsed === "object") savedByNamePrev = parsed;
                }
            } catch (eLBN) {}

            progress.close();
            var picked = confirmSharedSourceMode(sharedGroupsList, sharedSourceModeDefault, savedByNamePrev);
            progress = makeProgressPanel();

            // Always persist the captured selections (cancel OR continue) so
            // the dialog opens with the same picks next time.
            try {
                // Merge with whatever was previously saved so picks for
                // sources NOT in this run's selection survive.
                for (var mK in picked.modesByName) {
                    if (picked.modesByName.hasOwnProperty(mK)) {
                        savedByNamePrev[mK] = picked.modesByName[mK];
                    }
                }
                srSave("sharedSourceModesByName", JSON.stringify(savedByNamePrev));
            } catch (eSBN) {}
            // Also persist the global default (most-recent first pick) so
            // brand-new sources without a per-name entry land on the
            // user's most-used mode.
            for (var ssK in picked.modes) {
                if (picked.modes.hasOwnProperty(ssK)) {
                    try { srSave("sharedSourceMode", picked.modes[ssK]); } catch (eSSM) {}
                    break;
                }
            }

            if (picked.cancelled) {
                reportCancellation("Cancelled at the shared-source preflight \u2014 no roundtrip performed.");
                return;
            }
            sharedSourceModes = picked.modes;
            progress.update("Shared-source modes set\u2026", "", 11);
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

        // Each row: { cols:[shot, frames, res, notice, source, os, bake], layerIdx }
        // layerIdx=-1 for skip rows (shot comp already exists).
        var confRows = [];
        var confTotalFrames = 0;
        var confNameMaxLen   = 0;
        var confFramesMaxLen = 0;
        var confResMaxLen    = 0;
        var confNoticeMaxLen = 0;
        var confPathMaxLen   = 0;
        var confBakeMaxLen   = 0;

        // Resolves whether a given AVLayer is in the user's bake-marked
        // set. Linear scan — at most a handful of entries in practice.
        function isLayerBakeMarked(l) {
            for (var bki = 0; bki < revBakeLayers.length; bki++) {
                if (revBakeLayers[bki].layer === l) return true;
            }
            return false;
        }

        progress.update("Preparing Confirm Shots dialog\u2026", "0 of " + expandedLayers.length, 12);
        try {
        for (var cfi = 0; cfi < expandedLayers.length; cfi++) {
            if (cancelCheck()) return;
            // Per-row update so a hang/crash surfaces the exact offending row.
            progress.update(null, "row " + (cfi + 1) + " of " + expandedLayers.length, 12 + (2 * cfi / Math.max(1, expandedLayers.length)));
            var cfNum  = autoShotNumbers ? autoShotNumbers[cfi] : (startNum + cfi * increment);
            var cfName = shotPrefix + pad(cfNum, 3);
            var cfItem = expandedLayers[cfi];

            // Already exists?
            if (!!confExisting[cfName + "_comp"]) {
                confRows.push({ cols: [cfName, "", "", "\u2014 skip", "already exists", "", ""], layerIdx: -1 });
                if (cfName.length > confNameMaxLen) confNameMaxLen = cfName.length;
                continue;
            }

            // Cut duration
            // Math.abs: for a reversed-stretch layer AE reports outPoint < inPoint
            // (the stretch negativity swaps the edges). The preview shows
            // pre-conversion state, so normalize here — duration is the same
            // either way, just unsigned.
            var cfCutFrames = Math.round(Math.abs(cfItem.layer.outPoint - cfItem.layer.inPoint) * cfFps);
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

            var cfFrames = "" + cfCutFrames;
            var cfRes    = (cfSrcW > 0) ? (cfSrcW + "\u00d7" + cfSrcH) : "";

            // Notice column: fps mismatch + time-effect tags + shared-source
            // tag. Reversed effects are already surfaced via the loud dialog
            // at the top of the run; we repeat them here per-shot so it's
            // clear which shots the warning covered.
            var cfNoticeParts = [];
            if (cfFpsMismatch) cfNoticeParts.push("fps " + cfSrcFps + "\u2260" + cfFps);
            var cfEffects = scanSelLayerForEffects(cfItem.layer);
            for (var cfe = 0; cfe < cfEffects.length; cfe++) {
                cfNoticeParts.push(cfEffects[cfe].label);
            }
            if (cfItem.originalSourceId && sharedGroups[cfItem.originalSourceId]) {
                cfNoticeParts.push("[shared source]");
            }
            var cfNotice = cfNoticeParts.join("  ");

            // Bake column: summarize the user's per-layer Bake choices from
            // the reversal warning dialog for this shot. Empty when the
            // shot has no reversed clips. "N/M" where N = baked count and
            // M = total reversed in this shot's tree.
            var cfBakeBaked = 0, cfBakeTotal = 0;
            for (var cfb = 0; cfb < cfEffects.length; cfb++) {
                if (!cfEffects[cfb].reversed) continue;
                cfBakeTotal++;
                if (isLayerBakeMarked(cfEffects[cfb].layer)) cfBakeBaked++;
            }
            var cfBake;
            if (cfBakeTotal === 0)                      cfBake = "";
            else if (cfBakeBaked === cfBakeTotal)       cfBake = "\u2713 " + cfBakeBaked + "/" + cfBakeTotal;
            else if (cfBakeBaked === 0)                 cfBake = "\u2014 " + cfBakeBaked + "/" + cfBakeTotal;
            else                                         cfBake = cfBakeBaked + "/" + cfBakeTotal;

            var cfOsMark = cfItem.overscan ? "\u2715" : "";
            confRows.push({ cols: [cfName, cfFrames, cfRes, cfNotice, cfPath, cfOsMark, cfBake], layerIdx: cfi });
            if (cfName.length   > confNameMaxLen)   confNameMaxLen   = cfName.length;
            if (cfFrames.length > confFramesMaxLen) confFramesMaxLen = cfFrames.length;
            if (cfRes.length    > confResMaxLen)    confResMaxLen    = cfRes.length;
            if (cfNotice.length > confNoticeMaxLen) confNoticeMaxLen = cfNotice.length;
            if (cfPath.length   > confPathMaxLen)   confPathMaxLen   = cfPath.length;
            if (cfBake.length   > confBakeMaxLen)   confBakeMaxLen   = cfBake.length;
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
        var confColOs     = 120; // width fits "Overscan Toggle" header
        var confColBake   = Math.max(confBakeMaxLen   * 8 + 24, 70); // "Bake" column
        var confColSource = Math.max(confPathMaxLen   * 8 + 24, 260);
        // Source column gets whatever space remains after the fixed columns.
        var confFixedW = confColShot + confColFrames + confColRes + confColNotice + confColOs + confColBake + 60;
        confColSource  = Math.min(confColSource, maxDlgW - confFixedW);
        confColSource  = Math.max(confColSource, 260);
        var confDlgW   = Math.min(confFixedW + confColSource, maxDlgW);
        confDlgW = Math.max(confDlgW, 800);

        var confDlg = new Window("dialog", "Confirm Shots");
        confDlg.orientation = "column"; confDlg.alignChildren = ["fill", "top"];
        confDlg.spacing = 8; confDlg.margins = 14;

        // Header info
        var confProjName = proj.file ? proj.file.displayName.replace(".aep", "") : "unsaved project";
        var confInfoTxt = confProjName + "   \u2022   " + mainComp.name + "   \u2022   " + cfFps + " fps   \u2022   handles: " + handleFrames;
        if (overscanPercent > 0) confInfoTxt += "   \u2022   +" + overscanPercent + "% overscan (toggle per shot with \u2715)";
        confDlg.add("statictext", undefined, confInfoTxt);

        if (chkSkipRender.value) {
            confDlg.add("statictext", undefined, "SKIP RENDER IS ON \u2014 comps will be built but nothing will be rendered or imported");
        }

        confDlg.add("statictext", undefined, expandedLayers.length + " shot" + (expandedLayers.length !== 1 ? "s" : "") + " will be created:");

        var shotsPnl = confDlg.add("panel", undefined, "Shots");
        shotsPnl.orientation = "column"; shotsPnl.alignChildren = ["fill", "top"];
        shotsPnl.margins = [10, 12, 10, 10]; shotsPnl.spacing = 4;

        // Sort controls. Each confRows[i] owns its layerIdx independent of
        // row position, so sorting doesn't break the overscan-toggle handler
        // (which looks up confRows[selectedListboxIndex].layerIdx).
        var confSortKey = -1;  // -1 = no sort applied (original shot-order)
        var confSortDir = 1;
        function confSortRows() {
            if (confSortKey < 0) return;
            confRows.sort(function (a, b) {
                var av = a.cols[confSortKey], bv = b.cols[confSortKey];
                // Frames and Res columns parse to numbers when possible so
                // "1000" sorts after "500" instead of alphabetically.
                if (confSortKey === 1 || confSortKey === 2) {
                    var an = parseFloat(av), bn = parseFloat(bv);
                    if (!isNaN(an) && !isNaN(bn)) {
                        if (an !== bn) return (an < bn ? -1 : 1) * confSortDir;
                    }
                }
                av = String(av || "").toLowerCase(); bv = String(bv || "").toLowerCase();
                if (av !== bv) return (av < bv ? -1 : 1) * confSortDir;
                return 0;
            });
        }
        var confSortRow = shotsPnl.add("group");
        confSortRow.orientation = "row"; confSortRow.alignChildren = ["left", "center"];
        confSortRow.spacing = 4;
        confSortRow.add("statictext", undefined, "Sort:");
        var CONF_LABELS = ["Shot", "Frames", "Res", "Notice", "Source", "Overscan", "Bake"];
        var confBtns = [];
        for (var cs = 0; cs < CONF_LABELS.length; cs++) {
            var cb = confSortRow.add("button", undefined, CONF_LABELS[cs]);
            cb.preferredSize = [86, 22];
            cb.onClick = (function (col) {
                return function () {
                    if (confSortKey === col) confSortDir = -confSortDir;
                    else { confSortKey = col; confSortDir = 1; }
                    confSortRows();
                    confRepopulate();
                    confRefreshBtns();
                };
            })(cs);
            confBtns.push(cb);
        }
        function confRefreshBtns() {
            for (var i = 0; i < confBtns.length; i++) {
                var arrow = (i === confSortKey) ? (confSortDir === 1 ? "  ↓" : "  ↑") : "";
                confBtns[i].text = CONF_LABELS[i] + arrow;
            }
        }

        var confLB = shotsPnl.add("listbox", undefined, [], {
            multiselect: true,
            numberOfColumns: 7,
            showHeaders: true,
            columnTitles: ["Shot", "Frames", "Res", "Notice", "Source", "Overscan Toggle", "Bake"],
            columnWidths: [confColShot, confColFrames, confColRes, confColNotice, confColSource, confColOs, confColBake]
        });
        var confLBH = Math.max(confRows.length * 22 + 40, 200);
        confLBH = Math.min(confLBH, 600);
        confLB.preferredSize = [confDlgW, confLBH];
        function confRepopulate() {
            try { confLB.removeAll(); } catch (eRA) {}
            for (var cfi = 0; cfi < confRows.length; cfi++) {
                var cfRow = confLB.add("item", confRows[cfi].cols[0]);
                cfRow.subItems[0].text = confRows[cfi].cols[1]; // frames
                cfRow.subItems[1].text = confRows[cfi].cols[2]; // res
                cfRow.subItems[2].text = confRows[cfi].cols[3]; // notice
                cfRow.subItems[3].text = confRows[cfi].cols[4]; // source
                cfRow.subItems[4].text = confRows[cfi].cols[5]; // os
                cfRow.subItems[5].text = confRows[cfi].cols[6]; // bake
            }
        }
        confRepopulate();
        confRefreshBtns();

        // Shared-source warning panel. Only shown when cross-selection
        // sharing is detected (>= 2 entries with the same originalSourceId
        // coming from different mainLayerIdx). Lists the groups so the
        // user can audit the selection before committing — each shared
        // source becomes N separate shots via the chain-duplication pre-
        // pass, each rendering only its visible portion.
        // Compact recap of the per-clip shared-source choices picked in
        // the dedicated preflight dialog above. Read-only — change by
        // re-running the roundtrip and choosing again at that dialog.
        if (sharedGroupsList.length > 0) {
            var sharedRecapCounts = { separate: 0, shared: 0, single: 0 };
            for (var srci = 0; srci < sharedGroupsList.length; srci++) {
                var sm = sharedSourceModes[sharedGroupsList[srci].originalSourceId] || "separate";
                if (sharedRecapCounts.hasOwnProperty(sm)) sharedRecapCounts[sm]++;
            }
            var sharedRecapParts = [];
            if (sharedRecapCounts.separate > 0) sharedRecapParts.push(sharedRecapCounts.separate + " separate");
            if (sharedRecapCounts.shared   > 0) sharedRecapParts.push(sharedRecapCounts.shared   + " shared");
            if (sharedRecapCounts.single   > 0) sharedRecapParts.push(sharedRecapCounts.single   + " single");
            confDlg.add("statictext", undefined,
                "Shared sources: " + sharedRecapParts.join(", ") + ".");
        }

        // Footer + toggle button
        var confFooterTxt = "Total: " + confTotalFrames + "  (" + (Math.round(confTotalFrames / cfFps * 10) / 10) + "s)   handles: " + handleFrames;
        confDlg.add("statictext", undefined, confFooterTxt);

        var confBtnGrp = confDlg.add("group");
        confBtnGrp.orientation = "row"; confBtnGrp.alignment = ["fill", "bottom"];
        var confSpacer  = confBtnGrp.add("statictext", undefined, ""); confSpacer.alignment = ["fill", "center"];
        var confToggleOs = confBtnGrp.add("button", undefined, "Toggle Overscan"); confToggleOs.preferredSize = [130, 28];
        var confCancel  = confBtnGrp.add("button", undefined, "Cancel");  confCancel.preferredSize  = [80,  28];
        var confOk      = confBtnGrp.add("button", undefined, "Process"); confOk.preferredSize      = [110, 28];

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
        if (confResult !== 1) {
            reportCancellation("Cancelled on the Confirm Shots preflight \u2014 no roundtrip performed.");
            return;
        }

        // sharedSourceMode was already captured + persisted at the
        // shared-source preflight dialog above (which fires before
        // Confirm Shots). Nothing to capture here.

        // ── Deferred Save-As + mutations (post-accept) ───────────────────────
        // Everything from here on modifies state. Save-As to the next _v##
        // version FIRST so every change lands in the new file and the
        // original sits on disk as the rollback point. If the Save-As
        // fails (disk full, permissions, etc.), bail out without touching
        // anything — the user's project is still the original on disk.
        versionedFile = saveAsNextVersion();
        if (!versionedFile) { progress.close(); return; }
        cancelStats.mutationsStarted = true;
        progress.update(
            "Working in new version: " + versionedFile.displayName,
            "Original is preserved on disk as the backup.",
            15
        );

        // Resolve the Roundtrip root folder up front for the post-render
        // bake step. The main loop further down re-resolves it as
        // `fsShots` against the same UI field; both yield the same path.
        var bakeShotsPathText = etShotsFolder.text;
        var bakeAepFolder = proj.file.parent;
        var fsBakeRoot = /^(\/|[A-Za-z]:)/.test(bakeShotsPathText)
                       ? new Folder(bakeShotsPathText)
                       : new Folder(bakeAepFolder.fsName + "/" + bakeShotsPathText);
        if (!fsBakeRoot.exists) fsBakeRoot.create();
        if (!fsBakeRoot.exists) { alert("Could not create roundtrip folder:\n" + fsBakeRoot.fsName); progress.close(); return; }

        // Capture the source FootageItem of every reversed layer the
        // user marked "Bake" in the warning dialog, RIGHT NOW — before
        // conversion + auto-precompose may invalidate the layer
        // references. The post-render bake step (inside the import loop
        // below) walks each shot's `_comp` tree for any of these sources
        // and, when found, duplicates the auto-rendered plate inside
        // `_stack`, applies stretch=-100, renders the stack out as a
        // reversed variant, and places the result back into the stack.
        var bakeSources = [];
        if (revBakeLayers && revBakeLayers.length > 0) {
            for (var bli = 0; bli < revBakeLayers.length; bli++) {
                var bl = revBakeLayers[bli].layer;
                if (!bl) continue;
                try {
                    if (bl.source) bakeSources.push(bl.source);
                } catch (eBSC) {}
            }
        }

        // Now do the deferred DOM mutations: rewrite stretch→remap and
        // auto-precompose time-remapped layers. Running these earlier
        // (e.g. right after the reversal warning's Continue button) meant
        // the project was already mutated if the user later cancelled on
        // Confirm Shots. Deferring to here gives a clean rollback all the
        // way up to the Process click.
        //
        // Snapshot per-shot overscan decisions before we rebuild expandedLayers
        // so the user's toggles in the Confirm Shots dialog survive. Shot
        // order and count are preserved across convert + auto-precompose
        // (conversion is in-place; auto-precompose wraps each top-level
        // time-remapped layer 1:1 into a container precomp), so index→shot
        // mapping stays stable.
        var overscanByIndex = [];
        for (var osi = 0; osi < expandedLayers.length; osi++) {
            overscanByIndex.push(!!expandedLayers[osi].overscan);
        }

        var postConvertSel = applyTimeEffectConversions();
        if (postConvertSel === null) { progress.close(); return; }
        selLayers = postConvertSel;

        // Flip baked layers' time-remap from descending to ascending in
        // mainComp.
        //
        // applyTimeEffectConversions just rewrote any negative-stretch
        // reversals into descending time-remap keys so the rest of the
        // pipeline can handle them. For layers the user marked Bake the
        // reversal is captured in the separate `_reversed.mov` rendered
        // into _stack later — having mainComp ALSO play the source
        // backward via remap would double-direction the edit. So for
        // each baked layer we keep the keys and their times intact (so
        // the cut + handle structure is preserved) but mirror the VALUES
        // across the cut centre, which turns descending into ascending →
        // forward playback over the same source range.
        if (bakeSources && bakeSources.length > 0) {
            (function reverseBakedLayersInTree(comp, visited) {
                if (!comp || !(comp instanceof CompItem)) return;
                visited = visited || {};
                var ck = "c" + (comp.id || comp.name);
                if (visited[ck]) return;
                visited[ck] = true;
                for (var li = 1; li <= comp.numLayers; li++) {
                    var L;
                    try { L = comp.layer(li); } catch (eL1) { continue; }
                    if (!L) continue;
                    try {
                        var srcMatch = false;
                        for (var bsi = 0; bsi < bakeSources.length; bsi++) {
                            if (L.source === bakeSources[bsi]) { srcMatch = true; break; }
                        }
                        if (srcMatch && L.timeRemapEnabled) {
                            var tr = L.property("Time Remap");
                            var nKeys = tr.numKeys;
                            if (nKeys >= 2) {
                                // Snapshot values, then write mirrored.
                                var vals = [];
                                for (var k = 1; k <= nKeys; k++) {
                                    try { vals.push(tr.keyValue(k)); }
                                    catch (eKV) { vals.push(0); }
                                }
                                for (var w = 1; w <= nKeys; w++) {
                                    try { tr.setValueAtKey(w, vals[nKeys - w]); } catch (eSV) {}
                                }
                                // Force linear so playback is constant-
                                // speed forward, matching the bake.
                                for (var iw = 1; iw <= nKeys; iw++) {
                                    try { tr.setInterpolationTypeAtKey(iw, KeyframeInterpolationType.LINEAR, KeyframeInterpolationType.LINEAR); } catch (eIK) {}
                                }
                            }
                        }
                        if (L.source instanceof CompItem) {
                            reverseBakedLayersInTree(L.source, visited);
                        }
                    } catch (eL2) {}
                }
            })(mainComp);
        }

        // Rebuild expandedLayers from the post-conversion selection. The
        // inner references (found.footageLayer, found.footageComp) on the
        // old expandedLayers are fine for precomp selections, but a top-
        // level direct-footage layer that got wrapped by autoPrecomposeTrimmed
        // now has a different mainComp layer reference — the fresh walk
        // picks up the new container layer correctly.
        progress.update("Re-expanding selection after conversion\u2026", "0 of " + selLayers.length, 20);
        expandedLayers = buildExpandedLayers(selLayers, skippedLayers, 20, 2, true);
        if (expandedLayers === null) return;
        for (var osj = 0; osj < expandedLayers.length; osj++) {
            expandedLayers[osj].overscan = (osj < overscanByIndex.length) ? overscanByIndex[osj] : false;
        }

        // Cross-selection shared-source handling. Three modes from the
        // dedicated shared-source preflight dialog (only shown when
        // sharing was actually detected during expansion):
        //
        //   separate  every reference's chain is deep-duplicated up front
        //             so each reference renders its own isolated shot.
        //             Runs BEFORE the main loop so the first reference's
        //             processing doesn't pre-mutate any shared precomp.
        //
        //   shared    references collapse into ONE shotComp that covers
        //             the UNION of all visible ranges in source time.
        //             First reference processes with the union widened
        //             range; subsequent references are flagged
        //             `.shareSkip` and the main loop skips them. The
        //             shared precomp's inner footage layer (post-
        //             replaceSource) covers every reference's cut, so
        //             nothing gets truncated.
        //
        //   single    only the FIRST reference becomes a shot. To keep
        //             the original source intact for the OTHER (skipped)
        //             references in mainComp, the first reference's
        //             chain is also deep-duplicated. Subsequent
        //             references are flagged `.singleSkip` and the
        //             main loop leaves them untouched.
        var xselSeenSrc    = {};
        var shareFirstIdx  = {}; // origId → first expandedLayers index
        var shareUnion     = {}; // origId → {start, end} in source time
        for (var xpi = 0; xpi < expandedLayers.length; xpi++) {
            var xpItem = expandedLayers[xpi];
            if (!xpItem.isPrecomp) continue;
            if (!xpItem.originalSourceId) continue;
            var xpOid   = xpItem.originalSourceId;
            var xpPrior = xselSeenSrc[xpOid];
            if (xpPrior === undefined) {
                xselSeenSrc[xpOid]   = xpItem.mainLayerIdx;
                shareFirstIdx[xpOid] = xpi;
                if (xpItem.found) {
                    shareUnion[xpOid] = {
                        start: xpItem.found.rangeStart,
                        end:   xpItem.found.rangeEnd
                    };
                }
                continue;
            }
            if (xpPrior === xpItem.mainLayerIdx) continue; // intra-selection, allowed
            var xpGroupMode = sharedSourceModes[xpOid] || "separate";
            if (xpGroupMode === "separate") {
                duplicateChainForItem(xpItem);
            } else if (xpGroupMode === "shared") {
                if (xpItem.found && shareUnion[xpOid]) {
                    if (xpItem.found.rangeStart < shareUnion[xpOid].start) {
                        shareUnion[xpOid].start = xpItem.found.rangeStart;
                    }
                    if (xpItem.found.rangeEnd > shareUnion[xpOid].end) {
                        shareUnion[xpOid].end = xpItem.found.rangeEnd;
                    }
                }
                xpItem.shareSkip = true;
            } else { // "single"
                xpItem.singleSkip = true;
            }
        }
        // Per-group post-pass. SHARED groups: widen the first reference's
        // range to the union and flag the shared inner precomp for the
        // shot-aware rename so other references that still point at it
        // read sensibly. SINGLE groups: duplicate the first reference's
        // chain too — otherwise its mutations leak to the original
        // source that the unprocessed sibling references still
        // reference. SEPARATE groups need no post-pass beyond the
        // per-entry duplication already done above.
        for (var xpu in shareFirstIdx) {
            if (!shareFirstIdx.hasOwnProperty(xpu)) continue;
            var xpuMode = sharedSourceModes[xpu] || "separate";
            var xpFirst = expandedLayers[shareFirstIdx[xpu]];
            if (xpuMode === "shared") {
                var xpU = shareUnion[xpu];
                if (xpFirst && xpFirst.found && xpU) {
                    xpFirst.found.rangeStart = xpU.start;
                    xpFirst.found.rangeEnd   = xpU.end;
                }
                // Only flag the inner-rename when the shared inner precomp
                // differs from the outer; the existing single-shot rename
                // already covers the outer case (turns it into
                // "shot_NN_container").
                if (xpFirst && xpFirst.found && xpFirst.found.footageComp &&
                    xpFirst.layer && xpFirst.found.footageComp !== xpFirst.layer.source) {
                    xpFirst.sharedRenameInner = true;
                }
            } else if (xpuMode === "single") {
                if (xpFirst) duplicateChainForItem(xpFirst);
            }
        }

        var aepFolder = proj.file.parent;
        // Accept either an absolute path (e.g. picked via the Browse button)
        // or a path relative to the .aep's parent folder (legacy default
        // "../Roundtrip"). Absolute paths start with "/" on Unix or "D:" on
        // Windows; everything else is joined against aepFolder.
        var shotsPathText = etShotsFolder.text;
        var fsShots = /^(\/|[A-Za-z]:)/.test(shotsPathText)
                    ? new Folder(shotsPathText)
                    : new Folder(aepFolder.fsName + "/" + shotsPathText);
        if (!fsShots.exists) fsShots.create();
        if (!fsShots.exists) { alert("Could not create shots folder:\n" + fsShots.fsName); progress.close(); return; }
        var fsScripts = fsShots;

        // Scaffold the Roundtrip/ and _grade/ README.txt files so the
        // handoff tree is self-documenting from day one. See
        // lib/write_readmes.jsx for the content. Both writes are non-fatal
        // (README is a nicety, not a requirement).
        try {
            var readmeHelper = new File((new File($.fileName)).parent.parent.fsName + "/lib/write_readmes.jsx");
            if (readmeHelper.exists) $.evalFile(readmeHelper);
            if (typeof writeRoundtripReadme === "function") writeRoundtripReadme(fsShots);
            var fsGrades = new Folder(fsShots.fsName + "/_grade");
            if (!fsGrades.exists) fsGrades.create();
            if (typeof writeGradeReadme === "function") writeGradeReadme(fsGrades);
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
            var intermedCompRegistry    = {}; // intermediate comp id → {firstNum, lastNum, firstShotName} (nested-wrapper rename)

            // Pre-build set of existing comp names so the already-processed check is O(1)
            // rather than scanning all project items for every shot.
            var existingCompNames = {};
            for (var ei = 1; ei <= proj.numItems; ei++) {
                if (proj.item(ei) instanceof CompItem) existingCompNames[proj.item(ei).name] = true;
            }

            // Clear the AE render queue so leftover items from prior runs
            // (partial renders, user-added comps, previous roundtrip passes)
            // don't co-render with this run's plates. Iterate backward so
            // the indices stay valid as we remove.
            try {
                for (var rqi = proj.renderQueue.numItems; rqi >= 1; rqi--) {
                    try { proj.renderQueue.item(rqi).remove(); } catch (eRqr) {}
                }
            } catch (eRqClr) {}

            for (var i = 0; i < expandedLayers.length; i++) {
                if (cancelCheck()) return;
                var item      = expandedLayers[i];
                var layer     = item.layer;
                var isPrecomp = item.isPrecomp;
                var currentNum = autoShotNumbers ? autoShotNumbers[i] : (startNum + (i * increment));
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
                    // Cross-selection sharing: the pre-pass above duplicated
                    // chains (separate / single modes) and/or marked
                    // entries with `.shareSkip` (shared mode — collapse
                    // into one shotComp covering the union) or
                    // `.singleSkip` (single mode — only first reference
                    // becomes a shot, others stay untouched). Skip both
                    // skip flags here.
                    if (item.shareSkip || item.singleSkip) continue;
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
                cancelStats.shotsCreated++;

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
                // Defensive: cutStart > safeDuration (clamped rawStart against a
                // very short source) makes the line above negative, which then
                // places the "cut out" marker BEFORE "cut in". Clamp to zero so
                // the markers collapse to the same instant instead of crossing.
                if (cutDuration < 0) cutDuration = 0;

                // Work area = EDITORIAL CUT only (not cut+handles) so RAM
                // preview and UI focus land on the visible cut. The render
                // queue uses explicit timeSpanStart/Duration below, so the
                // render range still covers cut+handles independently.
                shotComp.workAreaStart    = cutStart;
                shotComp.workAreaDuration = cutDuration;
                shotComp.markerProperty.setValueAtTime(cutStart, cutMarker("cut in"));
                shotComp.markerProperty.setValueAtTime(cutStart + cutDuration, cutMarker("cut out"));

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
                        // Capture original timing AND label colour before
                        // precompose invalidates the ref — AE's native
                        // precompose does not reliably preserve in/out OR
                        // the layer's label on the new mainComp wrapper,
                        // so we restore them ourselves.
                        var origInPointPC  = layer.inPoint;
                        var origOutPointPC = layer.outPoint;
                        var origLabelPC    = 0;
                        try { origLabelPC = layer.label; } catch (eLblPC0) {}
                        var layerIdx = layer.index; // read before precompose invalidates the ref
                        try { mainComp.layers.precompose([layerIdx], shotName + "_container" + osSuffix, true); } catch(ePC) {}
                        // Always re-fetch outside the try — whether precompose succeeded or failed,
                        // mainComp.layer(idx) is valid: it's the wrapper on success, the original on failure.
                        layer = mainComp.layer(layerIdx);
                        // Restore original in/out + label so the edit's cut
                        // placement and colour-coding survives.
                        try {
                            layer.startTime = 0;
                            layer.inPoint   = origInPointPC;
                            layer.outPoint  = origOutPointPC;
                        } catch (eTimingPC) {}
                        try { layer.label = origLabelPC; } catch (eLblPC1) {}
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

                    // Intermediate precomp rename: every wrapper precomp BETWEEN
                    // the outer selection's source and the raw footage would
                    // otherwise keep its original auto-generated name
                    // ("C2456.mov Comp 1") with no indication of which shot
                    // lives there. Walk the chain from innermost to outermost
                    // and rename each layer along the way.
                    //
                    // Naming convention (suffix reflects distance from the
                    // footage, so names read intuitively top-down):
                    //   depth 0 (immediate parent of footage): _inner
                    //   depth 1 (one level up):                _inner2
                    //   depth 2:                               _inner3
                    //   …
                    //
                    // Registry is keyed per wrapper-comp id using the same
                    // last-write-wins range pattern as the outer range bin:
                    // if the same wrapper hosts multiple shots it ends up as
                    // "shot_030_050_inner" rather than flip-flopping per shot.
                    //
                    // Never touches the outer container (it has its own naming
                    // via precompLayerRegistry + the "_container" suffix).
                    try {
                        var iChain = (item.found && item.found.layerChain) ? item.found.layerChain : [];
                        var iFComp = item.found && item.found.footageComp;
                        if (iFComp && iFComp !== item.layer.source) {
                            // Build wrapperComps innermost-to-outermost.
                            // footageComp is the innermost wrapper (depth 0).
                            // Chain is outer-to-inner, so its last entry's
                            // source === footageComp; step back to grab the
                            // outer wrappers.
                            var wrapperComps = [iFComp];
                            for (var wi = iChain.length - 2; wi >= 0; wi--) {
                                try { wrapperComps.push(iChain[wi].source); } catch(eWS) {}
                            }
                            for (var wci = 0; wci < wrapperComps.length; wci++) {
                                var wComp = wrapperComps[wci];
                                if (!wComp) continue;
                                // Skip shared comps: if this wrapper is referenced
                                // from anywhere outside the chain we just walked
                                // (e.g. another layer in mainComp also uses it,
                                // or it lives in another precomp the user keeps),
                                // renaming would surprise them by retitling user-
                                // meaningful work as a shot artifact. Pure throw-
                                // away wrappers like AE-auto "X.mov Comp 1" have
                                // usedIn.length === 1 (only their immediate parent
                                // in this chain) and still get renamed.
                                var wUsedCount = 1;
                                try { wUsedCount = wComp.usedIn.length; } catch(eUI) {}
                                if (wUsedCount > 1) continue;
                                var wcId;
                                try { wcId = wComp.id; } catch(eWcId) { continue; }
                                if (!intermedCompRegistry[wcId]) {
                                    intermedCompRegistry[wcId] = {
                                        firstNum:      currentNum,
                                        firstShotName: shotName,
                                        lastNum:       currentNum
                                    };
                                } else {
                                    intermedCompRegistry[wcId].lastNum = currentNum;
                                }
                                var wReg = intermedCompRegistry[wcId];
                                var suffix = (wci === 0) ? "_inner" : ("_inner" + (wci + 1));
                                var icName = (wReg.firstNum === wReg.lastNum)
                                           ? wReg.firstShotName + suffix
                                           : wReg.firstShotName + "_" + pad(wReg.lastNum, 3) + suffix;
                                try { wComp.name = icName + osSuffix; } catch(eIR) {}
                                // File into the shot's bin so it doesn't float
                                // at the project root. Single-shot wrappers go
                                // under /Shots/{shot}/; multi-shot range wrappers
                                // under /Shots/{shot}_{lastNum}/.
                                try {
                                    var icBinName = (wReg.firstNum === wReg.lastNum)
                                                  ? wReg.firstShotName
                                                  : wReg.firstShotName + "_" + pad(wReg.lastNum, 3);
                                    var icBin = getShotBin(binShots, icBinName);
                                    if (icBin) wComp.parentFolder = icBin;
                                } catch(eICBin) {}
                            }
                        }
                    } catch(eIRB) {}

                    // No container comp. Mark source as processed, replace footage layer
                    // inside the deepest precomp with shotComp. All transforms/effects on
                    // footageLayer are preserved in place — nothing to transplant.
                    // Key on the ORIGINAL source id (pre-replaceSource) so later iterations
                    // can still detect the dupe after this replacement mutates the live source.
                    // Record mainLayerIdx so the dedup check above can distinguish
                    // within-selection repeats (allowed → N shots) from cross-
                    // selection sharing (deduped → 1 shot). Only write on first
                    // occurrence so mainLayerIdx isn't clobbered by later same-
                    // selection iterations.
                    if (item.originalSourceId && !processedSourceIds[item.originalSourceId]) {
                        processedSourceIds[item.originalSourceId] = { mainLayerIdx: item.mainLayerIdx, bin: shotBin };
                    }

                    plateInner.property("Marker").setValueAtTime(cutStart, cutMarker("cut in"));
                    plateInner.property("Marker").setValueAtTime(cutStart + cutDuration, cutMarker("cut out"));

                    footageLayer.replaceSource(shotComp, false);

                    // Shared-mode rename of the shared inner precomp (e.g.
                    // "Cloud Space slow") so OTHER mainComp references that
                    // still point at it (.shareSkip entries) read shot-aware
                    // instead of keeping the original name. Only applies in
                    // SHARED mode, when the inner is distinct from the outer.
                    if (item.sharedRenameInner && item.found && item.found.footageComp) {
                        try { item.found.footageComp.name = shotName + "_shared" + osSuffix; } catch (eSRN) {}
                    }

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

                    // Trim the footageLayer (now showing shotComp) to the EDITORIAL
                    // CUT range, not the cut+handles plate range. Handles are kept
                    // INSIDE the shotComp for external tools (Nuke, dynamicLink);
                    // exposing them on the outer layer's timeline makes adjacent
                    // shots overlap — the trailing handle of shot N covers the
                    // leading frames of shot N+1 in the same precomp and the edit
                    // stops reconstructing cleanly.
                    var flCutIn = 0, flCutOut = 0;
                    try {
                        var flStretch = (footageLayer.stretch !== 0) ? footageLayer.stretch : 100;
                        if (footageLayer.timeRemapEnabled) {
                            // Source→comp mapping is nonlinear for time-remapped layers.
                            // The footage layer's in/out are already the cut bounds in comp time.
                            flCutIn  = footageLayer.inPoint;
                            flCutOut = footageLayer.outPoint;
                        } else {
                            flCutIn  = footageLayer.startTime + cutStart * (flStretch / 100);
                            flCutOut = footageLayer.startTime + (cutStart + cutDuration) * (flStretch / 100);
                        }
                        footageLayer.inPoint  = flCutIn;
                        footageLayer.outPoint = flCutOut;
                        footageLayer.property("Marker").setValueAtTime(flCutIn,  cutMarker("cut in"));
                        footageLayer.property("Marker").setValueAtTime(flCutOut, cutMarker("cut out"));
                    } catch(eTrim) {}

                    var precompSrc = layer.source;
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
                    var origLayerLabel = 0;
                    try { origLayerLabel = layer.label; } catch (eLblDF0) {}
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
                    // Restore the original layer's label colour on the new
                    // wrapper layer in mainComp. precompose hands it a
                    // default label otherwise, which loses any colour-
                    // coding the user set up upstream (e.g. via the
                    // "Color Time-Reverse Layers" pass).
                    try { mainComp.layer(layerIdx).label = origLayerLabel; } catch (eLblDF1) {}

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

                    // Container duration = shotComp duration (= source length +
                    // 1 buffer frame), so the inner _comp layer can be shown at
                    // its FULL intrinsic length inside the container, not
                    // trimmed to the cut+handles range. Container time aligns
                    // 1:1 with shotComp time (= source time), which also makes
                    // cut markers sit at the source-time cutStart/cutEnd — same
                    // positions as the plateInner markers inside shotComp.
                    var shotCompDurSec = shotComp.duration;
                    try {
                        containerComp.duration = shotCompDurSec;
                        containerComp.displayStartTime = 0;
                        // Work area = editorial cut (same rationale as shotComp).
                        containerComp.workAreaStart    = cutStart;
                        containerComp.workAreaDuration = cutDuration;
                    } catch(eDur) {}

                    if (containerInner) {
                        containerInner.replaceSource(shotComp, false);
                        // startTime=0 lines container time up with shotComp
                        // (= source) time. Trim inPoint/outPoint to the clip +
                        // handles span so the layer in the container only
                        // shows the used range — outside that range it's just
                        // empty source that Nuke / Resolve / dynamicLink don't
                        // need. Container duration stays at shotComp.duration
                        // so source-time identity with mainComp still holds.
                        containerInner.startTime = 0;
                        containerInner.inPoint   = fullStart;
                        containerInner.outPoint  = fullStart + fullDurationSec;
                        // Markers in source/container time: cut_in at cutStart,
                        // cut_out at cutStart + cutDuration. containerComp,
                        // plateInner (inside shotComp), and containerInner all
                        // share the same time axis now.
                        containerComp.markerProperty.setValueAtTime(cutStart,               cutMarker("cut in"));
                        containerComp.markerProperty.setValueAtTime(cutStart + cutDuration, cutMarker("cut out"));
                        plateInner.property("Marker").setValueAtTime(cutStart,               cutMarker("cut in"));
                        plateInner.property("Marker").setValueAtTime(cutStart + cutDuration, cutMarker("cut out"));
                        containerInner.property("Marker").setValueAtTime(cutStart,               cutMarker("cut in"));
                        containerInner.property("Marker").setValueAtTime(cutStart + cutDuration, cutMarker("cut out"));

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

                    // Align the mainComp precomp layer with the source-aligned
                    // container timeline. Container time equals source time, so
                    // we want mainComp time origInPoint (cut_in) to show
                    // container time cutStart. precomp_layer.startTime = (origInPoint - cutStart).
                    // Edit placement (inPoint..outPoint) is preserved exactly
                    // at origInPoint..origOutPoint.
                    if (layer) {
                        try {
                            layer.startTime = origInPoint - cutStart;
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

                // Lock the raw plate AFTER all marker / timing / effect writes
                // (both branches above set markers on plateInner). Reimported
                // variants stack inside the {shot}_stack precomp created below,
                // not on this layer.
                try { plateInner.locked = true; } catch (ePL) {}

                // Create the {shot}_stack precomp as a layer at the TOP of
                // _comp. Sized to clip + handles (matches the reimported
                // plate.mov that lands inside it post-render), with
                // displayStartTime carrying the source-TC offset. Reimported
                // variants (grades, VFX, denoised plates) stack inside it.
                var stackComp = null;
                try {
                    var stackCompName = shotName + "_stack" + osSuffix;
                    stackComp = proj.items.addComp(
                        stackCompName, osWidth, osHeight, source.pixelAspect,
                        fullDurationSec, shotComp.frameRate
                    );
                    stackComp.displayStartTime = fullStart;
                    stackComp.label = 14; // Cyan — stack precomp
                    try { stackComp.parentFolder = shotBin; } catch (ePB) {}

                    // Cut in / cut out markers. setValueAtTime on comp.markerProperty
                    // uses the comp's LOCAL (0-based) time regardless of
                    // displayStartTime. The stack comp runs 0..fullDurationSec
                    // internally, so cut-in lands at handleSec and cut-out at
                    // handleSec+cutDuration.
                    var cutInLocal  = cutStart - fullStart;   // = handleSec
                    var cutOutLocal = cutInLocal + cutDuration;
                    try { stackComp.markerProperty.setValueAtTime(cutInLocal,  cutMarker("cut in"));  } catch (eMkIn)  {}
                    try { stackComp.markerProperty.setValueAtTime(cutOutLocal, cutMarker("cut out")); } catch (eMkOut) {}

                    // "Do not modify" notice — a locked red null pinned to
                    // the bottom of the stack so the warning is visible but
                    // doesn't clutter the time ruler.
                    try {
                        var noticeNullName    = "Managed by Gegenschuss Shot Roundtrip \u2014 do not modify";
                        var noticeNull        = stackComp.layers.addNull(stackComp.duration);
                        noticeNull.name       = noticeNullName;
                        // Also rename the underlying source \u2014 AE shows the
                        // SOURCE name in the timeline's Source Name column,
                        // which keeps reading "Null 3" / "Null 7" / etc.
                        // unless the source itself is renamed. This way
                        // the warning text shows up regardless of which
                        // column mode the user has on (Source Name or
                        // Layer Name).
                        try { if (noticeNull.source) noticeNull.source.name = noticeNullName; } catch (eNullSrc) {}
                        noticeNull.label      = 1;     // Red
                        noticeNull.guideLayer = true;  // never rendered
                        noticeNull.comment    = "Reimported variants (renders, grades) land here automatically. Any changes to this stack precomp will be overwritten by Shot Roundtrip.";
                        noticeNull.moveToEnd();
                        noticeNull.locked     = true;
                    } catch (eNull) {}

                    // Outer layer spans fullStart..fullStart+fullDurationSec so
                    // _comp-time = source-time identity still holds (render
                    // queue timeSpanStart, dynamicLink, cut markers rely on it).
                    var ppLayer = shotComp.layers.add(stackComp);
                    try { ppLayer.startTime = fullStart; } catch (eST) {}
                    try { ppLayer.inPoint   = fullStart; } catch (eIP) {}
                    try { ppLayer.outPoint  = fullStart + fullDurationSec; } catch (eOP) {}
                    try { ppLayer.moveToBeginning(); } catch (eMB) {}

                    // Cut in / cut out LAYER markers on the outer stack-comp
                    // layer so they show on its strip in _comp's timeline.
                    // setValueAtTime on a layer's Marker property takes COMP
                    // time (same semantics as the plate layer's markers above
                    // at lines ~2830) — NOT layer-relative time. Earlier
                    // builds passed (cutStart - fullStart) assuming layer
                    // time; that placed markers near the strip's end.
                    try { ppLayer.property("Marker").setValueAtTime(cutStart,               cutMarker("cut in"));  } catch (ePLMi) {}
                    try { ppLayer.property("Marker").setValueAtTime(cutStart + cutDuration, cutMarker("cut out")); } catch (ePLMo) {}

                    // "Re-render" Checkbox Control effect — user ticks this in
                    // the Effect Controls panel during editing to flag the
                    // shot for re-rendering on the next Re-render Plates run.
                    // Re-render Plates reads + auto-resets this after success.
                    try {
                        var rrEff = ppLayer.Effects.addProperty("ADBE Checkbox Control");
                        rrEff.name = "Re-render";
                    } catch (eRR) {}
                } catch (ePlatePC) {
                    reportError("PLATE-PRECOMP", ePlatePC, "Plate precomp creation failed for " + shotName);
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

                renderItems.push({ n: shotName, p: pPath, c: shotComp, s: fullStart, w: osWidth, h: osHeight, cs: cutStart, cd: cutDuration, bin: shotBin, rq: rq, pc: stackComp, src: source });

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
                    origH:                   source.height,
                    fps:                     safeFPS
                });
            }

        } catch(e) { reportError("PREP", e); try { progress.close(); } catch(eCP){} return; } finally { app.endUndoGroup(); }

        if (cancelCheck()) return;
        if (!chkSkipRender.value && renderItems.length > 0) {
            // Close the palette for the duration of the render: AE blocks the
            // script during renderQueue.render() so palette updates freeze
            // anyway, and AE's own render window takes over the UI — use its
            // cancel button if you need to stop the render mid-flight.
            //
            // Skip the render call entirely when we didn't queue anything this
            // run (every shot was a duplicate, etc.). Calling render() on a
            // queue that only holds leftover items from a previous failed run
            // can trigger AE's "Overwrite existing file?" modal, which often
            // appears behind the main window and looks like a UI freeze.
            progress.update("Saving project and handing off to AE's render queue\u2026", "AE's render window will take over now.", 60);
            progress.close();
            var renderException = null;
            try { proj.save(); proj.renderQueue.render(); }
            catch (eRender) { renderException = eRender; }

            // Inspect per-item statuses to distinguish three outcomes:
            //   USER_STOPPED → user hit Cancel in AE's render window
            //   ERR_STOPPED  → render errored out (codec, disk, etc.)
            //   DONE         → rendered cleanly
            // renderQueue.render() can throw OR return silently on cancel
            // depending on the AE version, so item.status is the canonical
            // source of truth. Only inspect items WE queued this run
            // (tracked via renderItems[i].rq) — leftover items from prior
            // sessions are ignored.
            var rDone = 0, rUserStopped = 0, rErrStopped = 0, rOther = 0;
            for (var rsI = 0; rsI < renderItems.length; rsI++) {
                var rqi = renderItems[rsI].rq;
                if (!rqi) { rOther++; continue; }
                try {
                    if      (rqi.status === RQItemStatus.DONE)         rDone++;
                    else if (rqi.status === RQItemStatus.USER_STOPPED) rUserStopped++;
                    else if (rqi.status === RQItemStatus.ERR_STOPPED)  rErrStopped++;
                    else rOther++;
                } catch (eSt) { rOther++; }
            }

            if (rUserStopped > 0) {
                // User hit Cancel in AE's render window. Route through the
                // shared cancellation summary and skip ALL post-render
                // steps — importing partial plates, dynamicLink, Nuke,
                // XML would all produce broken output referencing files
                // that don't exist. Re-running the roundtrip detects
                // existing shot comps as already-processed and skips
                // them, so the already-finished renders are preserved.
                cancelStats.rendersDone = rDone;
                reportCancellation("Render cancelled in AE's render window.\n"
                    + "Rendered " + rDone + " of " + renderItems.length + " plate(s) before cancel.\n\n"
                    + "Re-run the roundtrip to finish the remaining shots \u2014 "
                    + "the already-built shot comps will be detected as existing and skipped.");
                return;
            }

            if (rErrStopped > 0 || renderException) {
                var hint = "Check OM Template and disk space.";
                if (rErrStopped > 0) {
                    hint = rErrStopped + " plate(s) errored out during render. " + hint;
                }
                reportError("RENDER",
                    renderException || { message: "Render stopped with errors", line: "\u2014" },
                    hint);
                return;
            }

            progress = makeProgressPanel();
            progress.update("Render finished, importing plates\u2026", "", 65);
        } else if (!chkSkipRender.value && renderItems.length === 0) {
            progress.update("No new shots to render — all were already processed or skipped.", "", 60);
        }

        app.beginUndoGroup("Roundtrip Finish");
        try {
         if (chkBurninEnabled.value) {
            // Burnin Fields precomp — a single project-wide comp holding four
            // named text layers (Project, Production Company, Agency, Client).
            // The user edits Source Text on each layer to fill in the burnin
            // metadata. The CTRL in mainComp (below) reads these cross-comp
            // and composes `<project> / <company> / <agency> / <client> /
            // <shot> / f <frame>` (skipping any empty fields).
            var fieldsComp = null;
            try {
                for (var fi2 = 1; fi2 <= proj.numItems; fi2++) {
                    var it2 = proj.item(fi2);
                    if (it2 instanceof CompItem && it2.name === "Burnin Fields") { fieldsComp = it2; break; }
                }
                if (!fieldsComp) {
                    fieldsComp = proj.items.addComp("Burnin Fields", 800, 400, 1, mainComp.duration, mainComp.frameRate);
                    try { fieldsComp.displayStartTime = 0; } catch (eDS) {}
                    try { fieldsComp.label = 2; } catch (eLb) {}
                    try { fieldsComp.parentFolder = binShots; } catch (eBF) {}
                    try {
                        fieldsComp.comment = "Edit these four text layers to configure the burnin prefix. "
                                           + "Shot Roundtrip's Burnin CTRL in mainComp reads them live.";
                    } catch (eCm) {}
                }
                // Apply the dialog values as the authoritative Source Text
                // for each field — user's edits in the dialog win over
                // whatever's currently in the Burnin Fields comp, since the
                // dialog was pre-populated from the comp and they're the
                // last word. Added in reverse so stack order top→bottom
                // reads Project, Production Company, Agency, Client.
                var fieldSpec = [
                    { name: "Client",             text: etBurninClient.text,  y: 240 },
                    { name: "Agency",             text: etBurninAgency.text,  y: 180 },
                    { name: "Production Company", text: etBurninCompany.text, y: 120 },
                    { name: "Project",            text: etBurninProject.text, y:  60 }
                ];
                for (var fsi = 0; fsi < fieldSpec.length; fsi++) {
                    var spec = fieldSpec[fsi];
                    var fL = fieldsComp.layers.byName(spec.name);
                    if (!fL) {
                        fL = fieldsComp.layers.addText(spec.text);
                        fL.name = spec.name;
                    } else {
                        try {
                            fL.locked = false;
                            var existingDoc = fL.property("Source Text").value;
                            existingDoc.text = spec.text;
                            fL.property("Source Text").setValue(existingDoc);
                        } catch (eSetTxt) {}
                    }
                    try {
                        fL.locked = false;
                        var fDoc = fL.property("Source Text").value;
                        fDoc.fontSize = 36;
                        fDoc.fillColor = [1, 1, 1];
                        fDoc.justification = ParagraphJustification.LEFT_JUSTIFY;
                        fL.property("Source Text").setValue(fDoc);
                        fL.property("Anchor Point").setValue([0, 0]);
                        fL.position.setValue([40, spec.y]);
                    } catch (eFL) {}
                }
            } catch (eFC) {}

            // Zoned burnin — "Burnin" precomp sized to mainComp, layered on
            // top. Four text layers at the frame's corners:
            //   TL: PROJECT | PRODUCTION COMPANY           (top-left)
            //   TR: AGENCY | CLIENT                        (top-right)
            //   BL: TL F<N> @ <TC>                         (bottom-left)
            //   BM: <SHOT>                                 (bottom-middle)
            //   BR: SRC F<N> @ <TC>                        (bottom-right)
            // Plus two full-width semi-transparent strips at top / bottom
            // to keep the text readable over any footage.
            // All text uppercase, Courier New monospace, separator "|".
            // Fields read live from Burnin Fields; shot + TCs read live
            // from mainComp (layer detection at current time).
            try {
                var mainCompNameEsc = mainComp.name.replace(/'/g, "\\'");

                // Remove any old burnin artifacts before rebuilding.
                function purgeMainCompLayer(nm) {
                    var L = mainComp.layers.byName(nm);
                    if (L) { try { L.locked = false; L.remove(); } catch (eRmL) {} }
                }
                purgeMainCompLayer("Burnin CTRL");
                purgeMainCompLayer("Burnin CTRL BG");
                purgeMainCompLayer("Burnin");
                for (var pi = proj.numItems; pi >= 1; pi--) {
                    var pit = proj.item(pi);
                    if (pit instanceof CompItem && pit.name === "Burnin") {
                        try { pit.remove(); } catch (eRmP) {}
                    }
                }

                // Expression builders. Each returns a string joined with "\n"
                // (AE's expression parser choked on "\r" in prior builds).
                // Labels on all four field corners (reviewers on Vimeo may
                // not know the top-left / top-right convention). Shot name
                // gets its own center-bottom zone so the frame-at-a-glance
                // reads: project/company top, agency/client top, shot
                // dead-centre, timeline TC bottom-left, source TC bottom-
                // right. SRC / TL prefixes disambiguate the two TCs.
                function exprTopFields() {
                    return [
                        "var fc = null;",
                        "try { fc = comp('Burnin Fields'); } catch(e) {}",
                        "function readF(n) {",
                        "    try { return fc.layer(n).text.sourceText + ''; }",
                        "    catch(e) { return ''; }",
                        "}",
                        "var parts = [];",
                        "if (fc) {",
                        "    var p = readF('Project');            if (p) parts.push(p);",
                        "    var c = readF('Production Company'); if (c) parts.push('PRODUCTION COMPANY: ' + c);",
                        "}",
                        "parts.join(' | ').toUpperCase();"
                    ].join("\n");
                }
                function exprTopRight() {
                    return [
                        "var fc = null;",
                        "try { fc = comp('Burnin Fields'); } catch(e) {}",
                        "function readF(n) {",
                        "    try { return fc.layer(n).text.sourceText + ''; }",
                        "    catch(e) { return ''; }",
                        "}",
                        "var parts = [];",
                        "if (fc) {",
                        "    var a = readF('Agency'); if (a) parts.push('AGENCY: ' + a);",
                        "    var cl = readF('Client'); if (cl) parts.push('CLIENT: ' + cl);",
                        "}",
                        "parts.join(' | ').toUpperCase();"
                    ].join("\n");
                }
                function exprBottomLeft() {
                    return [
                        "var fps = 1 / thisComp.frameDuration;",
                        "var tlFrame = Math.floor(time / thisComp.frameDuration);",
                        "var tlTC = timeToTimecode(time, fps, false);",
                        "('TL F' + tlFrame + ' @ ' + tlTC).toUpperCase();"
                    ].join("\n");
                }
                function exprBottomCenter() {
                    return [
                        "var m = null;",
                        "try { m = comp('" + mainCompNameEsc + "'); } catch(e) {}",
                        "var out = '';",
                        "if (m) {",
                        "    var t = time;",
                        "    for (var i = 1; i <= m.numLayers; i++) {",
                        "        var L = m.layer(i);",
                        "        if (!L.enabled) continue;",
                        "        if (L.inPoint > t) continue;",
                        "        if (L.outPoint <= t) continue;",
                        "        try {",
                        "            var nm = L.source.name;",
                        "            var mm = nm.match(/^(.+?)_(container|comp)(_OS)?$/i);",
                        "            if (mm) { out = mm[1].toUpperCase(); break; }",
                        "        } catch(e) {}",
                        "    }",
                        "}",
                        "out;"
                    ].join("\n");
                }
                function exprBottomRight() {
                    return [
                        "var m = null;",
                        "try { m = comp('" + mainCompNameEsc + "'); } catch(e) {}",
                        "var out = '';",
                        "if (m) {",
                        "    var layerT = 0;",
                        "    var found = false;",
                        "    var t = time;",
                        "    for (var i = 1; i <= m.numLayers; i++) {",
                        "        var L = m.layer(i);",
                        "        if (!L.enabled) continue;",
                        "        if (L.inPoint > t) continue;",
                        "        if (L.outPoint <= t) continue;",
                        "        try {",
                        "            var nm = L.source.name;",
                        "            var mm = nm.match(/^(.+?)_(container|comp)(_OS)?$/i);",
                        "            if (mm) { layerT = t - L.inPoint; found = true; break; }",
                        "        } catch(e) {}",
                        "    }",
                        "    if (found) {",
                        "        var fps = 1 / thisComp.frameDuration;",
                        "        var shotFrame = Math.round(layerT / thisComp.frameDuration) + 1001;",
                        "        var srcTC = timeToTimecode(layerT, fps, false);",
                        "        out = ('SRC F' + shotFrame + ' @ ' + srcTC).toUpperCase();",
                        "    }",
                        "}",
                        "out;"
                    ].join("\n");
                }

                // Create the Burnin precomp at mainComp dimensions.
                var burninComp = proj.items.addComp(
                    "Burnin",
                    mainComp.width, mainComp.height, mainComp.pixelAspect,
                    mainComp.duration, mainComp.frameRate
                );
                try { burninComp.displayStartTime = mainComp.displayStartTime; } catch (eBDS) {}
                try { burninComp.label = 2; } catch (eBLb) {}
                try { burninComp.parentFolder = binShots; } catch (eBPF) {}

                // Per-text BG solids sized to hug each corner text (no more
                // full-width strips). Each corner block carries a measurement
                // dummy string so the BG reflects realistic content width;
                // sourceRectAtTime after dummy setValue gives the exact pixel
                // size. Font bold + 48pt @ 1080p for small-screen legibility.
                //
                // Leave ~80px (@1080p) at the very bottom for Vimeo's review-
                // player controls (scrubber + TC + fullscreen overlay the
                // lower ~60px during playback on vimeo.com).
                //
                // All pixel values scale off comp height so the burnin
                // stays visually consistent from SD through 4K/8K (4K ≈ 2×
                // scale, 8K ≈ 4× scale).
                var burninScale     = mainComp.height / 1080;
                var vimeoReservedPx = Math.round(80 * burninScale);
                var fieldProj    = "";
                var fieldCompany = "";
                var fieldAgency  = "";
                var fieldClient  = "";
                try { fieldProj    = etBurninProject.text || ""; } catch (eEF1) {}
                try { fieldCompany = etBurninCompany.text || ""; } catch (eEF2) {}
                try { fieldAgency  = etBurninAgency.text  || ""; } catch (eEF3) {}
                try { fieldClient  = etBurninClient.text  || ""; } catch (eEF4) {}
                function joinBurninLabelled(pairs) {
                    // pairs[i] = [label, value]. Empty label → no prefix.
                    var out = [];
                    for (var i = 0; i < pairs.length; i++) {
                        if (pairs[i][1]) {
                            var prefix = pairs[i][0] ? (pairs[i][0] + ": ") : "";
                            out.push((prefix + pairs[i][1]).toUpperCase());
                        }
                    }
                    return out.join(" | ");
                }
                var measureTL = joinBurninLabelled([["", fieldProj], ["Production Company", fieldCompany]]);
                var measureTR = joinBurninLabelled([["Agency", fieldAgency], ["Client", fieldClient]]);
                var measureBL = "TL F9999 @ 00:00:99:99";
                var measureBM = "SHOT_XXX_XXX";
                var measureBR = "SRC F1099 @ 00:00:99:99";

                // addCorner — add text + sized BG pair. Hybrid justification:
                //   - CENTER_JUSTIFY for center alignment (self-centres live
                //     text on the position regardless of width, so the SHOT
                //     zone looks right whether the live value is "SHOT_010"
                //     or "SHOT_010_030").
                //   - LEFT_JUSTIFY + manual position shift for both left
                //     and right alignment. AE's RIGHT_JUSTIFY on point
                //     text doesn't actually align the right edge at the
                //     layer origin — text ends up left-anchored and clips
                //     off the frame. Shifting position.x = edgeX - r.width
                //     puts the measured text's right edge at edgeX instead.
                // `edgeX` is the comp-space x that the text aligns to:
                //   align=left   → edgeX is the text's LEFT edge
                //   align=right  → edgeX is the text's RIGHT edge
                //   align=center → edgeX is the text's CENTRE
                function addCorner(name, expr, measureStr, edgeX, yPos, align) {
                    var L = burninComp.layers.addText(" ");
                    L.name = name;
                    L.label = 2;
                    var tp = L.property("Source Text");
                    tp.setValue(new TextDocument(measureStr && measureStr.length > 0 ? measureStr : " "));
                    var td = tp.value;
                    // Fixed 48pt across all resolutions — user's preferred
                    // size for HD and UHD alike (UHD on Vimeo gets scaled
                    // down in-browser, so larger source text ends up too
                    // heavy visually).
                    td.fontSize = 48;
                    // Modern bold monospace. Source Code Pro ships with
                    // Adobe Creative Cloud so it's guaranteed on every AE
                    // install (macOS + Windows). Fall back through Courier
                    // variants if the CC bundle ever goes missing.
                    try { td.font = "SourceCodePro-Bold"; } catch (eSc1) {
                        try { td.font = "SourceCodePro-Regular"; } catch (eSc2) {
                            try { td.font = "CourierNewPS-BoldMT"; } catch (eBnt) {
                                try { td.font = "CourierNewPSMT"; } catch (eFnt) {}
                            }
                        }
                    }
                    td.fillColor = [1, 1, 1];
                    td.justification =
                        (align === "center") ? ParagraphJustification.CENTER_JUSTIFY :
                                               ParagraphJustification.LEFT_JUSTIFY;
                    tp.setValue(td);
                    L.property("Anchor Point").setValue([0, 0]);
                    L.position.setValue([edgeX, yPos]);

                    // Measure the rendered dummy. For CENTER_JUSTIFY, rect.left
                    // is ~-width/2; for LEFT_JUSTIFY it's ~0. Shift layer
                    // position to land the text where we want.
                    var padX = Math.round(14 * burninScale), padY = Math.round(6 * burninScale);
                    var bgW = 400, bgH = 44, bgCX = edgeX, bgCY = yPos;
                    var finalX = edgeX;
                    try {
                        var r = L.sourceRectAtTime(L.startTime, false);
                        if (r.width > 0 && r.height > 0) {
                            if (align === "right") {
                                // Shift LEFT_JUSTIFY text left by its width
                                // so the right edge sits at edgeX.
                                finalX = edgeX - r.width;
                                L.position.setValue([finalX, yPos]);
                            }
                            bgW  = Math.ceil(r.width)  + padX * 2;
                            bgH  = Math.ceil(r.height) + padY * 2;
                            // Universal BG-centre formula — works whether
                            // rect.left is 0 (LEFT) or -width/2 (CENTER).
                            bgCX = Math.round(finalX + r.left + r.width / 2);
                            bgCY = Math.round(yPos   + r.top  + r.height / 2);
                        }
                    } catch (eMR) {}

                    // Swap in the live expression.
                    tp.expression = expr;
                    try { tp.expressionEnabled = true; } catch (eEE) {}

                    // Matching BG solid, sat just under the text.
                    var bg = burninComp.layers.addSolid([0, 0, 0], name + " BG", bgW, bgH, 1);
                    bg.position.setValue([bgCX, bgCY]);
                    try { bg.property("ADBE Transform Group").property("ADBE Opacity").setValue(50); } catch (eBgO) {}
                    try { bg.moveAfter(L); } catch (eMv) {}
                }

                var margin      = Math.round(40 * burninScale);
                // Baselines tuned for 48pt bold @1080p; scale with comp
                // height. Top gives ~10px clearance above ascender; bottom
                // sits ~80px clear of Vimeo controls.
                var baselineTop = Math.round(60 * burninScale);
                var baselineBot = mainComp.height - vimeoReservedPx - Math.round(20 * burninScale);

                addCorner("BR: SRC TC",     exprBottomRight(),  measureBR, mainComp.width - margin, baselineBot, "right");
                addCorner("BM: Shot",       exprBottomCenter(), measureBM, mainComp.width / 2,       baselineBot, "center");
                addCorner("BL: TL TC",      exprBottomLeft(),   measureBL, margin,                    baselineBot, "left");
                addCorner("TR: Agency",     exprTopRight(),     measureTR, mainComp.width - margin, baselineTop, "right");
                addCorner("TL: Project",    exprTopFields(),    measureTL, margin,                    baselineTop, "left");

                // Drop the Burnin precomp onto mainComp as the topmost
                // layer so it renders on top of everything.
                var burninLayer = mainComp.layers.add(burninComp);
                burninLayer.name  = "Burnin";
                burninLayer.label = 2;
                try { burninLayer.moveToBeginning(); } catch (eBMB) {}
            } catch (eBurnin) {
                reportError("BURNIN", eBurnin, "Burnin precomp creation failed.");
            }
         } // end chkBurninEnabled

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
                    // Keep the rendered plate in a dedicated plate/ sub-bin
                    // (mirrors the on-disk layout) so the shot bin's root
                    // stays clean — just comps + live footage, with all
                    // outputs (plate, render, grade) in their own sub-bins.
                    imp.parentFolder = getShotBin(ri.bin, "plate");
                    imp.label = 9; // Green — VFX return mov
                    var tComp = ri.c;
                    var tPrecomp = ri.pc;
                    if (tComp instanceof CompItem && tPrecomp instanceof CompItem) {
                        // Reimported {shot}_plate.mov lands INSIDE the
                        // {shot}_stack precomp at startTime=0 (local-0 of
                        // the precomp = source-TC fullStart via
                        // displayStartTime). layer.startTime is in LOCAL
                        // time; the render queue's timeSpanStart is the
                        // one that reads displayed time.
                        var nL = tPrecomp.layers.add(imp);
                        nL.startTime = 0;
                        nL.position.setValue([ri.w/2, ri.h/2]);
                        nL.label = 11;
                        // Cut markers in precomp time (precomp t=0 represents
                        // source-time fullStart, so cut_in = ri.cs - ri.s).
                        var cutInPP  = ri.cs - ri.s;
                        var cutOutPP = cutInPP + ri.cd;
                        try { nL.property("Marker").setValueAtTime(cutInPP,  cutMarker("cut in"));  } catch (eM1) {}
                        try { nL.property("Marker").setValueAtTime(cutOutPP, cutMarker("cut out")); } catch (eM2) {}

                        // Bake step (only when the user marked at least one
                        // reversed layer in this shot's tree for Bake):
                        // render `_stack` reversed through a small wrapper
                        // comp and add the result as a top-layer variant in
                        // `_stack`. The wrapper carries a descending time-
                        // remap on `_stack` (AE's reliable reversal — bare
                        // stretch=-100 needs paired startTime juggling that
                        // setting it via script doesn't do, which is why
                        // the previous attempt rendered blank).
                        var bakeMatchedThisShot = false;
                        if (bakeSources && bakeSources.length > 0) {
                            for (var bki = 0; bki < bakeSources.length; bki++) {
                                if (compTreeContainsSource(tComp, bakeSources[bki], {})) {
                                    bakeMatchedThisShot = true; break;
                                }
                            }
                        }
                        if (bakeMatchedThisShot) {
                            var bakedFolder = new Folder(fsBakeRoot.fsName + "/_baked");
                            if (!bakedFolder.exists) bakedFolder.create();
                            if (bakedFolder.exists) {
                                // Always write to {shot}_reversed.mov, overwriting any
                                // existing file from a previous run. Re-runs replace
                                // the bake in place so the user has a single canonical
                                // reversed file per shot rather than a pile of
                                // _v01/_v02 variants.
                                var revFile = new File(bakedFolder.fsName + "/" + ri.n + "_reversed.mov");
                                try { if (revFile.exists) revFile.remove(); } catch (eRevRm) {}

                                var stackFR = (tPrecomp.frameRate > 0) ? tPrecomp.frameRate : 25;
                                var stackFrameDur = 1.0 / stackFR;
                                // Pad the wrapper by 1 frame so the render's
                                // [0, _stack.duration] sits safely inside.
                                var wrapDur = tPrecomp.duration + stackFrameDur;
                                var wrapComp = null;
                                try {
                                    wrapComp = proj.items.addComp(
                                        "_bake_wrap_" + ri.n,
                                        tPrecomp.width,
                                        tPrecomp.height,
                                        tPrecomp.pixelAspect,
                                        wrapDur,
                                        stackFR
                                    );
                                } catch (eWrapAdd) {}

                                if (wrapComp) {
                                    var wrapL = null;
                                    try { wrapL = wrapComp.layers.add(tPrecomp); } catch (eWLA) {}
                                    if (wrapL) {
                                        try { wrapL.startTime = 0; } catch (eWST) {}
                                        // Enable time-remap and FLIP the
                                        // values on the auto-created keys
                                        // (AE creates 2 keys at the layer's
                                        // start/end with forward source-
                                        // time values). Reversing the
                                        // values gives a descending remap;
                                        // keys stay at the same times so
                                        // we don't have to fight AE's
                                        // boundary handling. setValueAtKey
                                        // is more reliable than removing +
                                        // re-adding.
                                        try {
                                            wrapL.timeRemapEnabled = true;
                                            var wrapTR = wrapL.property("Time Remap");
                                            if (wrapTR.numKeys >= 2) {
                                                try { wrapTR.setValueAtKey(1, tPrecomp.duration); } catch (eWK1) {}
                                                try { wrapTR.setValueAtKey(2, 0);                  } catch (eWK2) {}
                                                try { wrapTR.setInterpolationTypeAtKey(1, KeyframeInterpolationType.LINEAR, KeyframeInterpolationType.LINEAR); } catch (eWI1) {}
                                                try { wrapTR.setInterpolationTypeAtKey(2, KeyframeInterpolationType.LINEAR, KeyframeInterpolationType.LINEAR); } catch (eWI2) {}
                                            }
                                        } catch (eWRT) {}

                                        // Render the wrapper.
                                        var revRq = null;
                                        var revOk = false;
                                        try {
                                            revRq = proj.renderQueue.items.add(wrapComp);
                                            revRq.timeSpanStart    = 0;
                                            revRq.timeSpanDuration = Math.max(stackFrameDur, tPrecomp.duration);
                                            var revOM = revRq.outputModule(1);
                                            var revFoundT = false;
                                            for (var revT = 0; revT < revOM.templates.length; revT++) {
                                                if (revOM.templates[revT] === omTemplate) revFoundT = true;
                                            }
                                            if (revFoundT) revOM.applyTemplate(omTemplate);
                                            revOM.file = revFile;
                                            for (var qqi = 1; qqi <= proj.renderQueue.numItems; qqi++) {
                                                var qqit = proj.renderQueue.item(qqi);
                                                if (qqit !== revRq && qqit.status === RQItemStatus.QUEUED) {
                                                    try { qqit.render = false; } catch (eRRf) {}
                                                }
                                            }
                                            try { progress.close(); } catch (eRP1) {}
                                            try {
                                                proj.renderQueue.render();
                                                revOk = waitForFile(revFile, 20);
                                            } catch (eRevR) { revOk = false; }
                                            try { progress = makeProgressPanel(); } catch (eRP2) {}
                                        } catch (eRevSetup) { revOk = false; }
                                        try { if (revRq) revRq.remove(); } catch (eRRm) {}

                                        if (revOk) {
                                            var revFootage = null;
                                            try { revFootage = proj.importFile(new ImportOptions(revFile)); } catch (eRIm) {}
                                            if (revFootage) {
                                                try { revFootage.parentFolder = getBinFolder("_baked"); } catch (eBakedBin) {}
                                                try {
                                                    var revL = tPrecomp.layers.add(revFootage);
                                                    revL.startTime = 0;
                                                    try { revL.position.setValue([ri.w / 2, ri.h / 2]); } catch (eBPos) {}
                                                    revL.label = 8;
                                                    try { revL.comment = "Reversed variant of " + ri.n + "_plate (rendered from _stack via wrapper time-remap)."; } catch (eRComm) {}
                                                    try { revL.property("Marker").setValueAtTime(cutInPP,  cutMarker("cut in"));  } catch (eRBM1) {}
                                                    try { revL.property("Marker").setValueAtTime(cutOutPP, cutMarker("cut out")); } catch (eRBM2) {}
                                                } catch (eRevAdd) {}
                                            }
                                        }
                                    }
                                    try { wrapComp.remove(); } catch (eWRm) {}
                                }
                            }
                        }

                        // Always add the shot's ORIGINAL forward source
                        // (raw FootageItem) into _stack as a guide layer
                        // for visual checking — both with and without the
                        // Bake option. Aligned so source-time ri.s
                        // (= fullStart, the start of the rendered range)
                        // lands at _stack local time 0 — same axis as
                        // every other layer in the stack. Guide layers
                        // never contribute to any render. Moved BELOW
                        // every variant (just above the "Managed by
                        // Gegenschuss…" notice null) at the end of this
                        // block.
                        var origGuide = null;
                        if (ri.src) {
                            try {
                                origGuide = tPrecomp.layers.add(ri.src);
                                try { origGuide.startTime = -ri.s; } catch (eOGS) {}
                                try {
                                    origGuide.inPoint  = 0;
                                    origGuide.outPoint = tPrecomp.duration;
                                } catch (eOGIO) {}
                                try { origGuide.position.setValue([ri.w / 2, ri.h / 2]); } catch (eOGP) {}
                                try { origGuide.guideLayer = true; } catch (eOGG) {}
                                try { origGuide.label = 12; } catch (eOGL) {} // Brown — original guide
                                try { origGuide.comment = "Original forward source — guide layer (does not render)."; } catch (eOGC) {}
                                try { origGuide.property("Marker").setValueAtTime(cutInPP,  cutMarker("cut in"));  } catch (eOGM1) {}
                                try { origGuide.property("Marker").setValueAtTime(cutOutPP, cutMarker("cut out")); } catch (eOGM2) {}
                            } catch (eOGAdd) {}
                        }

                        // Layer order in _stack: variants on top (most
                        // recent first), then forward auto-plate, then
                        // the original-source guide layer, then the
                        // "Managed by…" notice null at the very bottom.
                        // Find the notice null by name (its ref isn't in
                        // scope here) and re-anchor.
                        var noticeNull = null;
                        for (var nni = 1; nni <= tPrecomp.numLayers; nni++) {
                            try {
                                if (tPrecomp.layer(nni).name.indexOf("Managed by Gegenschuss") === 0) {
                                    noticeNull = tPrecomp.layer(nni); break;
                                }
                            } catch (eNN) {}
                        }
                        if (noticeNull) {
                            try { noticeNull.locked = false; } catch (eNNu1) {}
                            try { noticeNull.moveToEnd(); } catch (eNNm) {}
                            try { noticeNull.locked = true; } catch (eNNu2) {}
                        }
                        if (origGuide) {
                            try {
                                if (noticeNull) origGuide.moveBefore(noticeNull);
                                else            origGuide.moveToEnd();
                            } catch (eOGMv) {}
                        }

                        // Topmost-only audio in _stack. Reimported variants
                        // (auto-plate, reversed bake, plate variants from
                        // Import Returns later) all carry the same audio
                        // track; without this we'd render the same audio
                        // 2-3× layered. Walk top-down and enable audio on
                        // the first audio-bearing non-guide AVLayer only;
                        // mute every other audio-bearing layer.
                        try {
                            var topAudioFound = false;
                            for (var ai = 1; ai <= tPrecomp.numLayers; ai++) {
                                var aL;
                                try { aL = tPrecomp.layer(ai); } catch (eAL0) { continue; }
                                if (!aL || !(aL instanceof AVLayer)) continue;
                                if (!aL.hasAudio) continue;
                                var enableAudio = (!aL.guideLayer && !topAudioFound);
                                try { aL.audioEnabled = enableAudio; } catch (eAE) {}
                                if (enableAudio) topAudioFound = true;
                            }
                        } catch (eAudPass) {}

                        // In _comp: disable every layer except the plate
                        // precomp itself and guide layers. Finds the precomp's
                        // outer layer by source identity to avoid disabling it.
                        var outerPPLayer = null;
                        for (var oxx = 1; oxx <= tComp.numLayers; oxx++) {
                            try { if (tComp.layer(oxx).source === tPrecomp) { outerPPLayer = tComp.layer(oxx); break; } } catch (eOP2) {}
                        }
                        for (var xx = 1; xx <= tComp.numLayers; xx++) {
                            var L = tComp.layer(xx);
                            if (L === outerPPLayer) continue;
                            if (L.guideLayer) continue;
                            try { L.enabled = false; L.audioEnabled = false; } catch (eEnD) {}
                        }
                        importedShots.push(ri.n);
                        count++;
                        cancelStats.rendersDone++;
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

            // ── Extend handles on time-remapped inner container layers ─
            // The direct-footage container path explicitly extends its
            // inner layer to cut + handles. The time-remapped path
            // (where autoPrecomposeTrimmed wrapped the layer earlier)
            // does not — the inner layer keeps its original cut-only
            // in/out, so the handle frames render BLACK in both the
            // _container preview in mainComp AND in the dynamicLink
            // wrapper. convertStretchReversalToRemap already placed
            // keyframes at the handle boundaries (preTime/postTime), so
            // we just need to pull the layer's inPoint/outPoint out to
            // span those outer keys.
            //
            // Runs LAST so dynamicLink built off the un-extended layer's
            // math stays valid; the wrapper then naturally inherits the
            // newly-visible handle content from `_container`.
            //
            // Only mutates layers with timeRemapEnabled AND at least two
            // keyframes — a no-op for default (non-time-remapped) inner
            // layers, which the existing direct-footage extension
            // already handles.
            try {
                progress.update("Extending handles on time-remapped layers…", "", 95);
                var dlContSeen = 0, dlContExtended = 0;
                for (var pii = 1; pii <= proj.numItems; pii++) {
                    var pIt = null;
                    try { pIt = proj.item(pii); } catch (ePI1) { continue; }
                    if (!pIt || !(pIt instanceof CompItem)) continue;
                    if (!/_container(?:_OS)?$/.test(pIt.name)) continue;
                    dlContSeen++;
                    for (var cli = 1; cli <= pIt.numLayers; cli++) {
                        var cL = null;
                        try { cL = pIt.layer(cli); } catch (eCL1) { continue; }
                        if (!cL) continue;
                        if (!cL.timeRemapEnabled) continue;
                        var trX = null;
                        try { trX = cL.property("Time Remap"); } catch (eTRX) {}
                        if (!trX || trX.numKeys < 2) continue;
                        var firstKey = trX.numKeys, lastKey = 1;
                        // Outermost keyframe times in LAYER time. Walk the
                        // keys to find min/max — they don't always come
                        // back in time order from numKeys iteration.
                        var minT = trX.keyTime(1), maxT = trX.keyTime(1);
                        for (var kk = 2; kk <= trX.numKeys; kk++) {
                            var kt = trX.keyTime(kk);
                            if (kt < minT) minT = kt;
                            if (kt > maxT) maxT = kt;
                        }
                        // Snap to comp duration so we never extend past
                        // the end of the host comp (AE clamps anyway, but
                        // explicit avoids any edge).
                        var compDurX = pIt.duration;
                        if (minT < 0)        minT = 0;
                        if (maxT > compDurX) maxT = compDurX;
                        if (maxT <= minT)    continue;
                        // Only widen — never contract. If the existing
                        // in/out already spans the keys (e.g. someone
                        // ran this twice, or a future direct-footage
                        // path does this earlier), leave it alone.
                        var widened = false;
                        try {
                            if (cL.inPoint > minT)  { cL.inPoint  = minT; widened = true; }
                            if (cL.outPoint < maxT) { cL.outPoint = maxT; widened = true; }
                        } catch (eIO) {}
                        if (widened) dlContExtended++;
                    }
                }
                progress.update("Extended handles on " + dlContExtended + " of " + dlContSeen + " container layer(s)", "", 96);
            } catch (eExtH) {
                // Non-fatal — handle extension is a polish step.
            }

            if (nukeDataList.length > 0) {
                if (chkCreateNuke.value) {
                    var nukeFileName = proj.file.name.replace(".aep", "") + "_Comp.nk";
                    var nukeFile = new File(fsScripts.fsName + "/" + nukeFileName);
                    writeNukeScript(nukeFile, nukeDataList, stats.fps);

                    for (var nsi = 0; nsi < nukeDataList.length; nsi++) {
                        var nsd = nukeDataList[nsi];
                        var shotNkFile = new File(fsScripts.fsName + "/" + nsd.name + "/" + nsd.name + ".nk");
                        writeNukeShotScript(shotNkFile, nsd);
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
                var pMiss = rpt.add("panel", undefined, "Missing Plates (" + missingPlates.length + ")");
                pMiss.orientation = "column"; pMiss.alignChildren = ["fill", "top"];
                pMiss.spacing = 3; pMiss.margins = [10, 15, 10, 10];
                for (var mi = 0; mi < missingPlates.length; mi++) { pMiss.add("statictext", undefined, missingPlates[mi]); }
            }

            if (clampedShots.length > 0) {
                var pClamp = rpt.add("panel", undefined, "Handles Clamped");
                pClamp.orientation = "column"; pClamp.alignChildren = ["fill", "top"];
                pClamp.spacing = 4; pClamp.margins = [10, 15, 10, 10];
                var maxC = Math.min(clampedShots.length, 8);
                for (var c = 0; c < maxC; c++) { pClamp.add("statictext", undefined, clampedShots[c]); }
                if (clampedShots.length > 8) pClamp.add("statictext", undefined, "...and " + (clampedShots.length - 8) + " more.");
            }

            if (skippedLayers.length > 0) {
                var pWarn = rpt.add("panel", undefined, "Ignored Layers");
                pWarn.orientation = "column"; pWarn.alignChildren = ["fill", "top"];
                pWarn.spacing = 4; pWarn.margins = [10, 15, 10, 10];
                var maxS = Math.min(skippedLayers.length, 5);
                for (var s = 0; s < maxS; s++) { pWarn.add("statictext", undefined, skippedLayers[s]); }
                if (skippedLayers.length > 5) pWarn.add("statictext", undefined, "...and " + (skippedLayers.length - 5) + " more.");
            }

            if (dynBuildFails.length > 0) {
                var pDlErr = rpt.add("panel", undefined, "dynamicLink Errors (" + dynBuildFails.length + ")");
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