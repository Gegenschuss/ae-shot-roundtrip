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

    // Scripts grouped into panels. Order within each group = order in the UI.
    var SCRIPTS = [
        // ── Precompose ────────────────────────────────────────────────────
        {
            group:  "Precompose",
            label:  "Precompose Trimmed",
            tip:    "Precompose selected layers (move all attributes, full-length comp), then trim the new precomp layer to the original in/out.",
            file:   "../helpers/precompose_trimmed.jsx"
        },
        {
            group:  "Precompose",
            label:  "Precomp to Guide Preview",
            tip:    "Precompose selected layers (leave attributes), scale parent layer 2×, mark inner footage as guide + scale 0.5×, then halve the precomp dimensions. Half-res working copy for faster previews.",
            file:   "../helpers/precomp_to_guide_preview.jsx"
        },
        {
            group:  "Precompose",
            label:  "Oversize Precomp 2×",
            tip:    "Precompose selected layers (leave attributes), double inner scale + position, then double the precomp dimensions. Parent-comp precomp layer stays at 100% scale, so the precomp shows 2× larger in the parent.",
            file:   "../helpers/oversize_precomp_2x.jsx"
        },
        {
            group:  "Precompose",
            label:  "Extend Precomp Handles",
            tip:    "Add 50 frames of handle at both in and out of the selected precomp, shifting layers to keep content aligned.",
            file:   "../helpers/extend_precomp_handles.jsx"
        },
        {
            group:  "Precompose",
            label:  "Copy Comp Markers to Precomp",
            tip:    "Copy all composition markers from the active comp into the selected precomp layer's source comp.",
            file:   "../helpers/copy_comp_markers_to_precomp.jsx"
        },

        // ── Layer ─────────────────────────────────────────────────────────
        {
            group:  "Layer",
            label:  "Mute All Audio",
            tip:    "Disable the audio switch on every selected layer and recursively on all layers inside any precomps they reference.",
            file:   "../helpers/mute_all_audio.jsx"
        },
        {
            group:  "Layer",
            label:  "Reverse Stretch → Remap",
            tip:    "For each selected layer with negative stretch, rewrite the reversal as an equivalent time remap (stretch=100 + two keyframes reproducing the reversed playback). Same logic Shot Roundtrip runs automatically — exposed here for standalone use.",
            file:   "../helpers/reverse_stretch_to_remap.jsx"
        },

        // ── Effects ───────────────────────────────────────────────────────
        {
            group:  "Effects",
            label:  "List Color Effects",
            tip:    "Scan the selected layer(s), comp(s), or active comp recursively for every color-modification effect (Apply LUT, Curves, Levels, Color Balance, Hue/Sat, Lumetri, …) and show them in a list. Toggle enabled state on the selection or enable/disable everything in one click.",
            file:   "../helpers/list_color_effects.jsx"
        },
        {
            group:  "Effects",
            label:  "Invert Transform Effect",
            tip:    "For each selected layer, add a second Distort > Transform effect named \"Transform (Inverse)\" whose properties expression-link to the original Transform and mathematically invert it (anchor↔position, rotation+skew negate, scale reciprocal). Animation tracks automatically. Applied in series with the source, the composite is the identity.",
            file:   "../helpers/invert_transform_effect.jsx"
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
            : new Window("palette", "Gegenschuss · Little Toolbox", undefined, { resizeable: true });

        win.orientation  = "column";
        win.alignChildren = ["fill", "top"];
        win.spacing      = 10;
        win.margins      = 12;

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

        if (win instanceof Window) {
            win.center();
            win.show();
        } else {
            win.layout.layout(true);
        }
    }

    buildUI(thisObj);

})(this);
