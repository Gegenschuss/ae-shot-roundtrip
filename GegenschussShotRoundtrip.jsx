/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * AE Shot Roundtrip Panel
 *
 * A dockable ScriptUI launcher for the AE Shot Roundtrip toolset.
 *
 * INSTALLATION (dockable panel):
 *   Copy this file to:
 *   /Applications/Adobe After Effects <version>/Scripts/ScriptUI Panels/
 *   Then find it under the Window menu in AE.
 *
 * QUICK RUN (floating window):
 *   File > Scripts > Run Script File… → pick this file.
 *   The panel and scripts must stay in the same "AE Shot Roundtrip" folder.
 */

(function (thisObj) {

    // Resolve the folder this panel lives in, then build paths to sibling scripts.
    var BASE = (new File($.fileName)).parent;

    var SCRIPTS = [
        {
            group:  "Roundtrip",
            label:  "Shot Roundtrip",
            tip:    "Extract selected shots for VFX work and re-import the renders back into the edit.",
            file:   "shot-roundtrip/shot_roundtrip.jsx"
        },
        {
            group:  "Roundtrip",
            label:  "Import Renders & Grades",
            tip:    "Scan all *_comp compositions for VFX renders (in each shot's render/ folder) and Resolve grades (in the flat _grade/ folder), and import them back into the project stacked grade > render > plate.",
            file:   "import-renders/import_renders.jsx"
        },
        {
            group:  "Roundtrip",
            label:  "Re-render Plates",
            tip:    "Re-render every shot's original plate to {shot}_{suffix}.mov (default suffix: denoised) ready for external denoising or stabilization, and import the results back into each _plate precomp as plate variants.",
            file:   "re-render-plates/re_render_plates.jsx"
        },
        {
            group:  "SynthEyes",
            label:  "Convert JSX \u2192 AEP",
            tip:    "Batch-run SynthEyes-exported .jsx files and save each result as a .aep project.",
            file:   "syntheyes-convert-jsx-to-aep/batch_syntheyes_to_ae.jsx"
        },
        {
            group:  "SynthEyes",
            label:  "Import AEP to AE",
            tip:    "Import SynthEyes .aep files and wire each one into its matching shot folder.",
            file:   "syntheyes-import-aep-to-ae/import_syntheyes_aeps_to_shots.jsx"
        },
        {
            group:  "Helpers",
            label:  "Export Fresh Shot XML",
            tip:    "Export an FCPXML 1.8 timeline from all *_comp compositions for import in DaVinci Resolve.",
            file:   "export-shot-xml/export_shot_xml.jsx"
        },
        {
            group:  "Helpers",
            label:  "Create dynamicLink Comps",
            tip:    "Standalone dynamicLink builder. Prompts for handle frames, then for each selected precomp or footage layer creates a wrapper comp with full cut+2\u00d7handles duration (black padded) in /Shots/dynamicLink.",
            file:   "create-dynamiclink-comps/create_dynamiclink_comps.jsx"
        }
    ];

    // ── helpers ────────────────────────────────────────────────────────────────

    function runScript(relativePath) {
        var f = new File(BASE.fsName + "/" + relativePath);
        if (!f.exists) {
            alert("Script not found:\n" + f.fsName);
            return;
        }
        $.evalFile(f);
    }

    // Collect unique group names in the order they first appear.
    function uniqueGroups() {
        var seen = {}, order = [];
        for (var i = 0; i < SCRIPTS.length; i++) {
            var g = SCRIPTS[i].group;
            if (!seen[g]) { seen[g] = true; order.push(g); }
        }
        return order;
    }

    // ── UI ─────────────────────────────────────────────────────────────────────

    function buildUI(thisObj) {
        var win = (thisObj instanceof Panel)
            ? thisObj
            : new Window("palette", "Shot Roundtrip", undefined, { resizeable: true });

        win.orientation  = "column";
        win.alignChildren = ["fill", "top"];
        win.spacing      = 12;
        win.margins      = 16;

        // ── logo ───────────────────────────────────────────────────────────────
        var logoRow = win.add("group");
        logoRow.orientation = "row";
        logoRow.alignment = ["center", "top"];
        logoRow.spacing = 10;
        logoRow.margins = [0, 4, 0, 4];

        var logoFile = new File(BASE.fsName + "/Gegenschuss.png");
        if (logoFile.exists) {
            var logoImg = logoRow.add("image", undefined, logoFile);
            logoImg.alignment = ["left", "center"];
        }
        var brandCol = logoRow.add("group");
        brandCol.orientation = "column";
        brandCol.alignment = ["left", "center"];
        brandCol.spacing = 2;
        var buildDate = brandCol.add("statictext", undefined, "20260419");
        // NOTE to editor: keep this date in sync with ship day (YYYYMMDD).
        buildDate.graphics.font = ScriptUI.newFont("Helvetica", "REGULAR", 11);
        buildDate.graphics.foregroundColor = buildDate.graphics.newPen(buildDate.graphics.PenType.SOLID_COLOR, [0.55, 0.55, 0.55, 1], 1);
        var ghLabel = brandCol.add("statictext", undefined, "github.com/");
        ghLabel.graphics.font = ScriptUI.newFont("Helvetica", "REGULAR", 10);
        ghLabel.graphics.foregroundColor = ghLabel.graphics.newPen(ghLabel.graphics.PenType.SOLID_COLOR, [0.4, 0.4, 0.4, 1], 1);
        var ghName = brandCol.add("statictext", undefined, "Gegenschuss");
        ghName.graphics.font = ScriptUI.newFont("Helvetica", "REGULAR", 10);
        ghName.graphics.foregroundColor = ghName.graphics.newPen(ghName.graphics.PenType.SOLID_COLOR, [0.4, 0.4, 0.4, 1], 1);

        var groups = uniqueGroups();

        for (var g = 0; g < groups.length; g++) {
            var groupName = groups[g];

            var panel = win.add("panel", undefined, groupName);
            panel.orientation  = "column";
            panel.alignChildren = ["fill", "top"];
            panel.spacing      = 6;
            panel.margins      = [10, 16, 10, 10];

            for (var s = 0; s < SCRIPTS.length; s++) {
                if (SCRIPTS[s].group !== groupName) continue;
                (function (script) {
                    var btn = panel.add("button", undefined, script.label);
                    btn.helpTip = script.tip;
                    btn.onClick = function () { runScript(script.file); };
                })(SCRIPTS[s]);
            }
        }

        // ── toolbox ────────────────────────────────────────────────────────────
        var closeRow = win.add("group");
        closeRow.orientation = "row";
        closeRow.alignment   = ["right", "bottom"];
        closeRow.margins     = [4, 0, 4, 0];
        var toolboxBtn = closeRow.add("button", undefined, "\u2692 Toolbox");
        toolboxBtn.preferredSize = [90, 22];
        toolboxBtn.helpTip = "Open Little Toolbox";
        toolboxBtn.onClick = function () { runScript("little-toolbox/GegenschussLittleToolbox.jsx"); };

        if (win instanceof Window) {
            win.center();
            win.show();
        } else {
            win.layout.layout(true);
        }
    }

    buildUI(thisObj);

})(this);
