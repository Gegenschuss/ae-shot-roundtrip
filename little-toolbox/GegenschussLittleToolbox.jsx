/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Little Toolbox Panel
 *
 * A dockable ScriptUI panel for general-purpose AE helper scripts.
 * Lives alongside the AE Shot Roundtrip panel but is independent of
 * the VFX roundtrip workflow.
 *
 * INSTALLATION (dockable panel):
 *   Copy this file to:
 *   /Applications/Adobe After Effects <version>/Scripts/ScriptUI Panels/
 *   Then find it under the Window menu in AE.
 *
 * QUICK RUN (floating window):
 *   File > Scripts > Run Script File… → pick this file.
 *   The panel and scripts must stay in the same folder.
 */

(function (thisObj) {

    var BASE = (new File($.fileName)).parent;

    var SCRIPTS = [
        {
            label:  "Mute All Audio",
            tip:    "Disable the audio switch on every selected layer and recursively on all layers inside any precomps they reference.",
            file:   "../helpers/mute_all_audio.jsx"
        },
        {
            label:  "Precomp to Guide Preview",
            tip:    "Precompose selected layers (leave attributes), scale parent layer 2\u00d7, mark inner footage as guide + scale 0.5\u00d7, then halve the precomp dimensions.",
            file:   "../helpers/precomp_to_guide_preview.jsx"
        },
        {
            label:  "Copy Comp Markers to Precomp",
            tip:    "Copy all composition markers from the active comp into the selected precomp layer's source comp.",
            file:   "../helpers/copy_comp_markers_to_precomp.jsx"
        },
        {
            label:  "Extend Precomp Handles",
            tip:    "Add 50 frames of handle at both in and out of the selected precomp, shifting layers to keep content aligned.",
            file:   "../helpers/extend_precomp_handles.jsx"
        },
        {
            label:  "Precompose Trimmed",
            tip:    "Precompose selected layers (move all attributes, full-length comp), then trim the new precomp layer to the original in/out.",
            file:   "../helpers/precompose_trimmed.jsx"
        },
        {
            label:  "Reverse Stretch \u2192 Remap",
            tip:    "For each selected layer with negative stretch, rewrite the reversal as an equivalent time remap (stretch=100 + two keyframes reproducing the reversed playback). Same logic Shot Roundtrip runs automatically \u2014 exposed here for standalone use.",
            file:   "../helpers/reverse_stretch_to_remap.jsx"
        },
        {
            label:  "List Color Effects",
            tip:    "Scan the selected comp(s) recursively for every color-modification effect (Apply LUT, Curves, Levels, Color Balance, Hue/Sat, Lumetri, …) and show them in a list. Toggle enabled state on the selection or enable/disable everything in one click.",
            file:   "../helpers/list_color_effects.jsx"
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

    // ── UI ─────────────────────────────────────────────────────────────────────

    function buildUI(thisObj) {
        var win = (thisObj instanceof Panel)
            ? thisObj
            : new Window("palette", "Gegenschuss \u00b7 Little Toolbox", undefined, { resizeable: true });

        win.orientation  = "column";
        win.alignChildren = ["fill", "top"];
        win.spacing      = 10;
        win.margins      = 12;

        // ── buttons ────────────────────────────────────────────────────────────
        for (var s = 0; s < SCRIPTS.length; s++) {
            (function (script) {
                var btn = win.add("button", undefined, script.label);
                btn.helpTip = script.tip;
                btn.onClick = function () { runScript(script.file); };
            })(SCRIPTS[s]);
        }

        if (win instanceof Window) {
            win.center();
            win.show();
        } else {
            win.layout.layout(true);
        }
    }

    buildUI(thisObj);

})(this);
