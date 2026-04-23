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
    // Each entry has a single-char Unicode icon (BMP range so AE's ScriptUI
    // font stack renders it on both macOS + Windows — no emojis) and a text
    // label. Tiny mode shows the icon only; normal mode shows "icon  label".
    var SCRIPTS = [
        // ── Precompose ────────────────────────────────────────────────────
        {
            group:  "Precompose",
            icon:   "▣",
            label:  "Precompose Trimmed",
            tip:    "Precompose selected layers (move all attributes, full-length comp), then trim the new precomp layer to the original in/out.",
            file:   "helpers/precompose_trimmed.jsx"
        },
        {
            group:  "Precompose",
            icon:   "▥",
            label:  "Precomp to Guide Preview",
            tip:    "Precompose selected layers (leave attributes), scale parent layer 2×, mark inner footage as guide + scale 0.5×, then halve the precomp dimensions. Half-res working copy for faster previews.",
            file:   "helpers/precomp_to_guide_preview.jsx"
        },
        {
            group:  "Precompose",
            icon:   "▦",
            label:  "Oversize Precomp 2×",
            tip:    "Precompose selected layers (leave attributes), double inner scale + position, then double the precomp dimensions. Parent-comp precomp layer stays at 100% scale, so the precomp shows 2× larger in the parent.",
            file:   "helpers/oversize_precomp_2x.jsx"
        },
        {
            group:  "Precompose",
            icon:   "↔",
            label:  "Extend Precomp Handles",
            tip:    "Add 50 frames of handle at both in and out of the selected precomp, shifting layers to keep content aligned.",
            file:   "helpers/extend_precomp_handles.jsx"
        },
        {
            group:  "Precompose",
            icon:   "⚑",
            label:  "Copy Comp Markers to Precomp",
            tip:    "Copy all composition markers from the active comp into the selected precomp layer's source comp.",
            file:   "helpers/copy_comp_markers_to_precomp.jsx"
        },

        // ── Layer ─────────────────────────────────────────────────────────
        {
            group:  "Layer",
            icon:   "♪",
            label:  "Mute All Audio",
            tip:    "Disable the audio switch on every selected layer and recursively on all layers inside any precomps they reference.",
            file:   "helpers/mute_all_audio.jsx"
        },
        {
            group:  "Layer",
            icon:   "⇄",
            label:  "Reverse Stretch → Remap",
            tip:    "For each selected layer with negative stretch, rewrite the reversal as an equivalent time remap (stretch=100 + two keyframes reproducing the reversed playback). Same logic Shot Roundtrip runs automatically — exposed here for standalone use.",
            file:   "helpers/reverse_stretch_to_remap.jsx"
        },

        // ── Effects ───────────────────────────────────────────────────────
        {
            group:  "Effects",
            icon:   "◉",
            label:  "List Color Effects",
            tip:    "Scan the selected layer(s), comp(s), or active comp recursively for every color-modification effect (Apply LUT, Curves, Levels, Color Balance, Hue/Sat, Lumetri, …) and show them in a list. Toggle enabled state on the selection or enable/disable everything in one click.",
            file:   "helpers/list_color_effects.jsx"
        },
        {
            group:  "Effects",
            icon:   "⇌",
            label:  "Invert Transform Effect",
            tip:    "For each selected layer, add a second Distort > Transform effect named \"Transform (Inverse)\" whose properties expression-link to the original Transform and mathematically invert it (anchor↔position, rotation+skew negate, scale reciprocal). Animation tracks automatically. Applied in series with the source, the composite is the identity.",
            file:   "helpers/invert_transform_effect.jsx"
        },

        // ── View ──────────────────────────────────────────────────────────
        {
            group:  "View",
            icon:   "⟦⟧",
            label:  "Work Area → Cut Markers",
            tip:    "Set the active comp's work area to span its \"cut in\" and \"cut out\" composition markers. Handy on _comp / _container comps after a roundtrip.",
            file:   "helpers/set_work_area_to_cut_markers.jsx"
        },

        // ── Project ───────────────────────────────────────────────────────
        // Pipeline / interchange helpers that operate on the whole project
        // or active comp (vs. the layer-scoped tools above). Moved out of
        // the main Shot Roundtrip panel so it stays focused on the
        // end-to-end flow; these are one-off grabs you reach for as needed.
        {
            group:  "Project",
            icon:   "↪",
            label:  "Export Shot XML",
            tip:    "Export an FCP7 XML for DaVinci Resolve. Two modes: \"Shots folder\" scans every *_comp under the Shots folder (one clip per shot, appended head-to-tail) — used by the full Shot Roundtrip. \"Active composition\" dumps every footage layer in the current comp at its own timeline position (one track per AE layer) — the export half of the Comp Grade Roundtrip, paired with Import Comp Grades.",
            file:   "export-shot-xml/export_shot_xml.jsx"
        },
        {
            group:  "Project",
            icon:   "↩",
            label:  "Import Comp Grades",
            tip:    "Import half of the Comp Grade Roundtrip. For each footage layer in the active comp, finds a matching graded file in the _grade/ folder (next to the AEP) and drops it in as a new layer directly above the original. Matches by source-file stem prefix — Resolve's \"Use Unique Filenames\" suffix (e.g. _V1-0064) is fine. Newest-by-modification-time wins if multiple versions are present. Aligns by embedded source timecode (QuickTime tmcd atom) so Resolve-rendered handles stack correctly; falls back to inPoint alignment when TC can't be read.",
            file:   "import-comp-grades/import_comp_grades.jsx"
        },
        {
            group:  "Project",
            icon:   "⟡",
            label:  "Create dynamicLink Comps",
            tip:    "Standalone dynamicLink builder. Prompts for handle frames, then for each selected precomp or footage layer creates a wrapper comp with full cut + 2× handles duration (black padded) in /Shots/dynamicLink.",
            file:   "create-dynamiclink-comps/create_dynamiclink_comps.jsx"
        },
        // ── Language ──────────────────────────────────────────────────────
        {
            group:  "Language",
            icon:   "◐",
            label:  "Toggle Language",
            tip:    "Project-wide language switch. Walks every comp in the project, detects layers tagged via layer.comment \"lang:XX\" or layer name suffix \"_lang_XX\" (case-insensitive — \"_LANG_DE\", \"Lang: en\" all work; XX = 2-8 letter/digit/underscore code), and enables the picked language while disabling siblings in other languages. Untagged layers are left alone. Mirrors the audio switch so dubbed audio swaps alongside visuals. Last-used language is remembered across sessions.",
            file:   "helpers/toggle_language.jsx"
        },
        {
            group:  "Language",
            icon:   "◎",
            label:  "Language Preflight",
            tip:    "Audit all language-tagged layers across the project before a batch render. Shows comp / layer / code / source (comment or name) / on-off state in a multi-column list, filterable by language. Reveal selected layer jumps to it in its comp. Copy as text dumps the audit to the clipboard. Read-only — never mutates the project.",
            file:   "helpers/language_preflight.jsx"
        },
        {
            group:  "Language",
            icon:   "⟳",
            label:  "Render Languages",
            tip:    "Batch-render one or more render queue items across every tagged language. Two targets: (1) AE Render Queue — blocking, in-place, duplicates the template per (template × language) pair with _<lang> suffix; (2) Adobe Media Encoder — saves a {project}_<lang>.aep per language and queueInAMEs each, so AME reads an immutable per-language file. Pick multiple templates; total renders = templates × languages. Restores the project's original state after the batch.",
            file:   "helpers/render_languages.jsx"
        }
    ];

    // ── Tiny-mode persistence ──────────────────────────────────────────────────
    var SETTINGS_SECTION = "Gegenschuss Little Toolbox";
    var SETTINGS_TINY    = "tiny";
    function loadTiny() {
        try {
            if (app.settings.haveSetting(SETTINGS_SECTION, SETTINGS_TINY)) {
                return app.settings.getSetting(SETTINGS_SECTION, SETTINGS_TINY) === "1";
            }
        } catch (e) {}
        return false;
    }
    function saveTiny(value) {
        try { app.settings.saveSetting(SETTINGS_SECTION, SETTINGS_TINY, value ? "1" : "0"); } catch (e) {}
    }

    function buttonLabel(script, tiny) {
        return tiny ? script.icon : (script.icon + "  " + script.label);
    }

    // Per-group foreground colors. Applied via button.graphics.foregroundColor —
    // ScriptUI may ignore this on macOS Aqua-styled buttons; if icons come out
    // grey everywhere, switch to iconbutton + PNG assets (not wired yet).
    var GROUP_COLORS = {
        "Precompose": [0.45, 0.75, 1.00, 1.0], // cyan-ish
        "Layer":      [0.55, 0.90, 0.55, 1.0], // green
        "Effects":    [1.00, 0.65, 0.35, 1.0], // orange
        "View":       [0.85, 0.55, 1.00, 1.0], // violet
        "Language":   [0.95, 0.55, 0.80, 1.0], // pink
        "Project":    [0.95, 0.85, 0.45, 1.0]  // gold
    };
    function paintButton(btn, groupName) {
        var c = GROUP_COLORS[groupName];
        if (!c) return;
        try {
            var pen = btn.graphics.newPen(btn.graphics.PenType.SOLID_COLOR, c, 1);
            btn.graphics.foregroundColor = pen;
        } catch (ePc) {}
    }

    function tinyBtnLabel(on) {
        return (on ? "◼" : "◻") + "  tiny";
    }

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
        win.spacing      = 6;
        win.margins      = 8;

        var tinyMode = loadTiny();

        // Button sizing for tiny mode. Tweak if icons need more breathing room.
        var TINY_BTN_W    = 40;
        var TINY_BTN_H    = 26;
        var TINY_SPACING  = 4;
        var tinyLastCols  = -1; // memo, skip rebuild if column count unchanged

        function hoverTip(script) {
            // Two-part tooltip: name on first line, blank line, description.
            return script.label + "\n\n" + script.tip;
        }

        // Builds the tiny-mode toggle in the given parent container. Returns
        // { btn, useIconBtn }. Used in two spots — normal mode places it in
        // a top row, tiny mode places it as the first cell of the flow grid.
        function makeTinyToggle(parent, sizeW, sizeH) {
            var btn;
            var useIconBtn = false;
            try {
                var logoFile = new File(BASE.fsName + "/Gegenschuss-tiny.png");
                if (logoFile.exists) {
                    var logoImg = ScriptUI.newImage(logoFile);
                    btn = parent.add("iconbutton", undefined, logoImg, { toggle: true, style: "toolbutton" });
                    btn.preferredSize = [sizeW, sizeH];
                    btn.value = tinyMode;
                    useIconBtn = true;
                }
            } catch (eLogo) {}
            if (!useIconBtn) {
                btn = parent.add("button", undefined, tinyBtnLabel(tinyMode));
                btn.preferredSize = [sizeW, sizeH];
            }
            btn.helpTip = "tiny mode\n\n"
                        + "Icon-only buttons in a wrapping flow layout. "
                        + "Resize the panel and the icons reflow across rows. "
                        + "Setting persists across sessions.";
            btn.onClick = function () {
                if (useIconBtn) tinyMode = !!btn.value;
                else {
                    tinyMode = !tinyMode;
                    btn.text = tinyBtnLabel(tinyMode);
                }
                saveTiny(tinyMode);
                applyMode(tinyMode);
            };
            return btn;
        }

        function destroyAllChildren() {
            while (win.children.length > 0) {
                try { win.remove(win.children[0]); } catch (e) {}
            }
            tinyLastCols = -1;
        }

        function buildNormal() {
            destroyAllChildren();

            // Top row: tiny toggle, right-aligned.
            var topRow = win.add("group");
            topRow.orientation = "row";
            topRow.alignChildren = ["right", "center"];
            topRow.alignment = ["fill", "top"];
            topRow.spacing = 4; topRow.margins = 0;
            makeTinyToggle(topRow, 28, 28);

            // Group panels stacked vertically.
            var content = win.add("group");
            content.orientation = "column";
            content.alignChildren = ["fill", "top"];
            content.alignment = ["fill", "fill"];
            content.spacing = 10; content.margins = 0;

            var groups = uniqueGroups();
            for (var g = 0; g < groups.length; g++) {
                var groupName = groups[g];
                var panel = content.add("panel", undefined, groupName);
                panel.orientation = "column";
                panel.alignChildren = ["fill", "top"];
                panel.spacing = 6;
                panel.margins = [10, 16, 10, 10];
                for (var s = 0; s < SCRIPTS.length; s++) {
                    if (SCRIPTS[s].group !== groupName) continue;
                    (function (script) {
                        var btn = panel.add("button", undefined, buttonLabel(script, false));
                        btn.helpTip = hoverTip(script);
                        paintButton(btn, script.group);
                        btn.onClick = function () { runScript(script.file); };
                    })(SCRIPTS[s]);
                }
            }
        }

        function buildTiny(availableWidth) {
            // Total cells = 1 toggle + N scripts. Compute how many fit per
            // row, rebuild only when that number changes.
            var totalCells = 1 + SCRIPTS.length;
            var avail = Math.max(TINY_BTN_W, availableWidth - 16);
            var cols  = Math.max(1, Math.floor((avail + TINY_SPACING) / (TINY_BTN_W + TINY_SPACING)));
            if (cols === tinyLastCols) return;
            tinyLastCols = cols;

            destroyAllChildren();
            tinyLastCols = cols;

            var content = win.add("group");
            content.orientation = "column";
            content.alignChildren = ["left", "top"];
            content.alignment = ["fill", "fill"];
            content.spacing = TINY_SPACING; content.margins = 0;

            for (var i = 0; i < totalCells; i += cols) {
                var row = content.add("group");
                row.orientation = "row";
                row.alignChildren = ["left", "center"];
                row.alignment = ["left", "top"];
                row.spacing = TINY_SPACING; row.margins = 0;
                for (var j = 0; j < cols && (i + j) < totalCells; j++) {
                    var cellIdx = i + j;
                    if (cellIdx === 0) {
                        // Toggle takes the leading cell.
                        makeTinyToggle(row, TINY_BTN_W, TINY_BTN_H);
                    } else {
                        (function (script) {
                            var btn = row.add("button", undefined, script.icon);
                            btn.helpTip = hoverTip(script);
                            btn.preferredSize = [TINY_BTN_W, TINY_BTN_H];
                            paintButton(btn, script.group);
                            btn.onClick = function () { runScript(script.file); };
                        })(SCRIPTS[cellIdx - 1]);
                    }
                }
            }
        }

        function panelWidth() {
            try {
                if (win.size && win.size[0]) return win.size[0];
                if (win.bounds)              return win.bounds.width;
            } catch (eW) {}
            return 300;
        }

        function applyMode(tiny) {
            if (tiny) buildTiny(panelWidth());
            else      buildNormal();
            try { win.layout.layout(true); win.layout.resize(); } catch (eL) {}
        }

        // onResizing fires during the drag; onResize fires when it stops.
        // Wiring both keeps the flow responsive + final.
        function onResize() {
            if (tinyMode) buildTiny(panelWidth());
            try { win.layout.layout(true); win.layout.resize(); } catch (eOR) {}
        }
        win.onResizing = onResize;
        win.onResize   = onResize;

        applyMode(tinyMode);

        if (win instanceof Window) {
            win.center();
            win.show();
        } else {
            win.layout.layout(true);
        }
    }

    buildUI(thisObj);

})(this);
