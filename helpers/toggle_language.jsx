/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Toggle Language
 *
 * Project-wide language switch for multi-language AE deliverables.
 * No expressions, no Essential Graphics — just mechanical enable /
 * disable of tagged layers.
 *
 * Tag layers one of two ways:
 *
 *   A) Layer.comment starts with "lang:<code>"
 *      e.g. comment  = "lang:de"
 *      e.g. comment  = "LANG: EN"
 *      e.g. comment  = "lang:en  (stronger serif variant)"
 *      Preferred — keeps the Timeline layer name clean. (Layer comments
 *      are a separate per-layer string, NOT markers — visible in the
 *      Timeline's "Comments" column, editable via right-click → Comment.)
 *
 *   B) Layer name ends with "_lang_<code>"
 *      e.g. name     = "Title_lang_de"
 *      e.g. name     = "Title_LANG_DE"
 *      e.g. name     = "VO_english_lang_en"
 *      Fallback — easier to scan visually if you have many layers.
 *
 * Both forms are CASE-INSENSITIVE — "_lang_", "_LANG_", "_Lang_" all
 * work; "de", "DE", "De" all normalize to "de" internally. <code> is
 * any 2–8 letter/digit/underscore string (ISO 639-1 "de"/"en", custom
 * "de_at", whatever you like — the script groups by the normalized
 * lowercase form).
 *
 * How it works:
 *   1. Walks every CompItem in the project (not just the active comp —
 *      a language swap should affect the whole deliverable).
 *   2. For each layer, reads its language tag (comment first, name second).
 *   3. Shows a dialog listing every detected code + layer count, with a
 *      dropdown to pick the active language (remembers your last pick
 *      across sessions).
 *   4. Enables every layer whose tag matches the active language,
 *      disables every layer whose tag is any OTHER language. Untagged
 *      layers are left alone. Audio-enabled flag is mirrored alongside
 *      the visibility flag so dubbed audio tracks swap too.
 *
 * Single undo step per run.
 */

(function () {

    var proj = app.project;

    // Grey ScriptUI confirmation — consistent with the rest of the toolkit.
    // Call instead of alert() for any end-of-run message or precondition error.
    function greyAlert(title, msg) {
        var dlg = new Window("dialog", title);
        dlg.orientation = "column"; dlg.alignChildren = ["fill", "top"];
        dlg.spacing = 10; dlg.margins = 14;
        var p = dlg.add("panel", undefined, "");
        p.orientation = "column"; p.alignChildren = ["fill", "top"];
        p.margins = [12, 12, 12, 12]; p.spacing = 4;
        var lines = String(msg).split("\n");
        for (var i = 0; i < lines.length; i++) {
            p.add("statictext", undefined, lines[i]);
        }
        var bg = dlg.add("group");
        bg.orientation = "row"; bg.alignment = ["fill", "bottom"];
        bg.add("statictext", undefined, "").alignment = ["fill", "center"];
        var ok = bg.add("button", undefined, "OK", { name: "ok" });
        ok.preferredSize = [90, 28];
        ok.onClick = function () { dlg.close(1); };
        dlg.show();
    }

    var NAME_RE    = /_lang_([a-z0-9_]{2,8})$/i;
    var COMMENT_RE = /^\s*lang\s*:\s*([a-z0-9_]{2,8})/i;

    function detectLang(layer) {
        try {
            if (layer.comment) {
                var m = COMMENT_RE.exec(layer.comment);
                if (m) return m[1].toLowerCase();
            }
        } catch (eC) {}
        var nm = "";
        try { nm = String(layer.name || ""); } catch (eN) {}
        var m2 = NAME_RE.exec(nm);
        return m2 ? m2[1].toLowerCase() : null;
    }

    // ── scan ───────────────────────────────────────────────────────────

    var comps = [];
    for (var i = 1; i <= proj.numItems; i++) {
        var it = proj.item(i);
        if (it instanceof CompItem) comps.push(it);
    }

    var found = {};       // code → { count, compsSet, layers: [{comp, layer}] }
    var totalTagged = 0;

    for (var c = 0; c < comps.length; c++) {
        var cm = comps[c];
        for (var li = 1; li <= cm.numLayers; li++) {
            var L = cm.layer(li);
            var code = detectLang(L);
            if (!code) continue;
            if (!found[code]) found[code] = { count: 0, compsSet: {}, layers: [] };
            found[code].count++;
            found[code].compsSet[cm.id] = true;
            found[code].layers.push({ comp: cm, layer: L });
            totalTagged++;
        }
    }

    var codes = [];
    for (var k in found) codes.push(k);
    codes.sort();

    if (codes.length === 0) {
        greyAlert("Toggle Language",
              "No language-tagged layers found.\n\n"
            + "Tag layers with either:\n"
            + "    layer.comment starting with \"lang:XX\", or\n"
            + "    layer name ending in \"_lang_XX\"\n\n"
            + "XX is any 2–8 letter code (e.g. de, en, fr, de_at).");
        return;
    }

    // ── dialog ─────────────────────────────────────────────────────────

    var SECTION = "Gegenschuss Toggle Language";
    var KEY     = "activeLang";
    var defaultCode = codes[0];
    try {
        if (app.settings.haveSetting(SECTION, KEY)) {
            var saved = app.settings.getSetting(SECTION, KEY);
            if (saved && found[saved]) defaultCode = saved;
        }
    } catch (eR) {}

    var dlg = new Window("dialog", "Toggle Language");
    dlg.orientation = "column"; dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 10; dlg.margins = 14;

    var pickPnl = dlg.add("panel", undefined, "Active language");
    pickPnl.orientation = "column"; pickPnl.alignChildren = ["fill", "top"];
    pickPnl.margins = [10, 15, 10, 10]; pickPnl.spacing = 6;

    var row = pickPnl.add("group");
    row.orientation = "row"; row.alignChildren = ["left", "center"]; row.spacing = 8;
    row.add("statictext", undefined, "Language:");
    var dd = row.add("dropdownlist", undefined, codes);
    var defaultIdx = -1;
    for (var di = 0; di < codes.length; di++) {
        if (codes[di] === defaultCode) { defaultIdx = di; break; }
    }
    if (defaultIdx >= 0) dd.selection = defaultIdx; else dd.selection = 0;

    var foundPnl = dlg.add("panel", undefined, "Tagged layers in project");
    foundPnl.orientation = "column"; foundPnl.alignChildren = ["fill", "top"];
    foundPnl.margins = [10, 15, 10, 10]; foundPnl.spacing = 2;
    for (var cI = 0; cI < codes.length; cI++) {
        var code = codes[cI];
        var compCount = 0;
        for (var _x in found[code].compsSet) compCount++;
        foundPnl.add("statictext", undefined,
            "  " + code + " — " + found[code].count + " layer(s) in " + compCount + " comp(s)");
    }
    var untagged = foundPnl.add("statictext", undefined, "Untagged layers are left alone.");
    untagged.graphics.font = ScriptUI.newFont("Helvetica", "ITALIC", 10);

    var btnGrp = dlg.add("group");
    btnGrp.orientation = "row"; btnGrp.alignment = ["fill", "bottom"];
    btnGrp.add("statictext", undefined, "").alignment = ["fill", "center"];
    var cancel = btnGrp.add("button", undefined, "Cancel", { name: "cancel" });
    cancel.preferredSize = [80, 28];
    var apply = btnGrp.add("button", undefined, "Apply", { name: "ok" });
    apply.preferredSize = [110, 28];

    cancel.onClick = function () { dlg.close(2); };
    apply.onClick  = function () {
        if (!dd.selection) return;
        dlg.close(1);
    };

    if (dlg.show() !== 1) return;
    var activeCode = dd.selection.text;

    // ── apply ──────────────────────────────────────────────────────────

    var enabledCnt  = 0;
    var disabledCnt = 0;

    app.beginUndoGroup("Toggle Language → " + activeCode);
    try {
        for (var key in found) {
            var shouldEnable = (key === activeCode);
            var entries = found[key].layers;
            for (var e = 0; e < entries.length; e++) {
                var layer = entries[e].layer;
                try { layer.enabled = shouldEnable; } catch (eE) {}
                // Mirror the audio switch so dubbed audio layers swap too.
                // hasAudio guards against shape/text layers etc.
                try {
                    if (layer.hasAudio) layer.audioEnabled = shouldEnable;
                } catch (eA) {}
                if (shouldEnable) enabledCnt++; else disabledCnt++;
            }
        }
    } finally {
        app.endUndoGroup();
    }

    try { app.settings.saveSetting(SECTION, KEY, activeCode); } catch (eS) {}

    greyAlert("Toggle Language → " + activeCode,
          "Enabled: "  + enabledCnt  + "\n"
        + "Disabled: " + disabledCnt);

})();
