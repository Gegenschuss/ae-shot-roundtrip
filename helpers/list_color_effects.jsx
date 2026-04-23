/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * List / Toggle Color-Modification Effects (recursive)
 *
 * Scans the selected comp(s) — or the active comp if none selected in the
 * Project panel — and walks every layer recursively into any precomp
 * source. Collects every color-modification effect (Apply Color LUT,
 * Curves, Levels, Color Balance, Hue/Saturation, Lumetri, …) and shows
 * them in a listbox. User can toggle enabled state on the selection or
 * bulk enable/disable everything. Shared precomps are visited once.
 *
 * Each toggle is wrapped in its own undo step so the user can back out
 * incrementally.
 */

(function () {

    var proj = app.project;
    if (!proj) { alert("Color Effects: open a project first."); return; }

    // ── Pick the starting set ────────────────────────────────────────────────
    // Priority order:
    //   1. Layers selected in the active comp's timeline → scan those
    //      layers + recurse into any precomp sources. Most targeted option,
    //      used when the artist is "in" a comp and only wants a subset.
    //   2. Comps selected in the Project panel → scan each fully.
    //   3. Active comp (no timeline selection) → scan the whole active
    //      comp.
    var startLayers = [];   // direct layers to inspect at top level
    var startComps  = [];   // comps to scan fully (recursively)
    var scanLabel   = "";
    var activeItem = proj.activeItem;
    if (activeItem instanceof CompItem) {
        try {
            var tlSel = activeItem.selectedLayers;
            if (tlSel && tlSel.length > 0) {
                for (var tl = 0; tl < tlSel.length; tl++) startLayers.push(tlSel[tl]);
                scanLabel = tlSel.length + " selected layer"
                          + (tlSel.length === 1 ? "" : "s") + " in " + activeItem.name;
            }
        } catch (eTLS) {}
    }
    if (startLayers.length === 0) {
        try {
            var psel = proj.selection;
            for (var s = 0; s < psel.length; s++) {
                if (psel[s] instanceof CompItem) startComps.push(psel[s]);
            }
        } catch (eSel) {}
        if (startComps.length > 0) {
            scanLabel = startComps.length + " selected comp"
                      + (startComps.length === 1 ? "" : "s");
        } else if (activeItem instanceof CompItem) {
            startComps.push(activeItem);
            scanLabel = "active comp " + activeItem.name;
        }
    }
    if (startLayers.length === 0 && startComps.length === 0) {
        alert("Color Effects: select layers in the active comp's timeline, "
            + "comps in the Project panel, or open a comp.");
        return;
    }

    // ── Known AE color-modification effect matchNames ────────────────────────
    // Stock effects from the Color Correction category. matchName is used
    // because it's stable across languages and user renames (unlike
    // effect.name which the user can edit).
    var COLOR_EFFECTS = {
        "ADBE Apply Color LUT2":        1,
        "ADBE Apply Color LUT":         1,
        "ADBE Auto Color":              1,
        "ADBE Auto Contrast":           1,
        "ADBE AutoLevels":              1,
        "ADBE Black&White":             1,
        "ADBE Brightness & Contrast 2": 1,
        "ADBE Broadcast Colors":        1,
        "ADBE CHANNEL MIXER":           1,
        "ADBE Change Color":            1,
        "ADBE Change To Color2":        1,
        "ADBE Color Balance 2":         1,
        "ADBE Color Balance (HLS)":     1,
        "ADBE Color Link":              1,
        "ADBE Color Stabilizer":        1,
        "ADBE Colorama":                1,
        "ADBE CurvesCustom":            1,
        "ADBE Easy Levels2":            1,
        "ADBE Equalize":                1,
        "ADBE Exposure2":               1,
        "ADBE Gamma/Pedestal/Gain2":    1,
        "ADBE HUE SATURATION":          1,
        "ADBE Leave Color":             1,
        "ADBE Lumetri":                 1,
        "ADBE Photo Filter":            1,
        "ADBE PROCAMP3":                1,
        "ADBE PROCAMP":                 1,
        "ADBE Pro Levels2":             1,
        "ADBE Selective Color":         1,
        "ADBE Shadow/Highlight":        1,
        "ADBE Tint":                    1,
        "ADBE Tritone":                 1,
        "ADBE Vibrance":                1
    };

    // ── Recurse + collect ────────────────────────────────────────────────────
    var entries = []; // { compName, layerName, effectName, effect }
    var visited = {}; // comp.id → true, so shared precomps are scanned once

    function scanLayer(layer, owningCompName) {
        // Collect color effects on this layer, then recurse into its
        // precomp source if there is one.
        try {
            var effects = layer.property("ADBE Effect Parade");
            if (effects && effects.numProperties > 0) {
                for (var e = 1; e <= effects.numProperties; e++) {
                    var fx = effects.property(e);
                    try {
                        if (COLOR_EFFECTS[fx.matchName]) {
                            entries.push({
                                compName:   owningCompName,
                                layerName:  layer.name,
                                effectName: fx.name,
                                effect:     fx
                            });
                        }
                    } catch (eFx) {}
                }
            }
        } catch (eEff) {}
        try {
            if (layer.source instanceof CompItem) scanComp(layer.source);
        } catch (eRec) {}
    }

    function scanComp(comp) {
        if (!(comp instanceof CompItem)) return;
        if (visited[comp.id]) return;
        visited[comp.id] = true;
        for (var i = 1; i <= comp.numLayers; i++) {
            scanLayer(comp.layer(i), comp.name);
        }
    }

    // Timeline selection → scan just those layers + recurse
    if (startLayers.length > 0 && activeItem instanceof CompItem) {
        for (var sL = 0; sL < startLayers.length; sL++) {
            scanLayer(startLayers[sL], activeItem.name);
        }
    }
    // Project-panel comps (or active comp fallback) → scan fully
    for (var sc = 0; sc < startComps.length; sc++) scanComp(startComps[sc]);

    if (entries.length === 0) {
        alert("Color Effects: no color-modification effects found in " + scanLabel + " (recursive).");
        return;
    }

    // Sort helpers — `sortKey` and `sortDir` govern the current order.
    // Columns: 0 = Comp, 1 = Layer, 2 = Effect, 3 = State. State sorts by
    // the live effect.enabled value so toggles + re-sort groups correctly.
    var SORT_KEYS = ["comp", "layer", "effect", "state"];
    var sortKey = 0;   // default: Comp
    var sortDir = 1;   // 1 = asc, -1 = desc
    function stateText(on) { return on ? "On" : "Off"; }
    function isEnabled(ent) {
        try { return !!ent.effect.enabled; } catch (eE) { return true; }
    }
    function sortEntries() {
        entries.sort(function (a, b) {
            var av, bv;
            if (sortKey === 0)      { av = a.compName;   bv = b.compName;   }
            else if (sortKey === 1) { av = a.layerName;  bv = b.layerName;  }
            else if (sortKey === 2) { av = a.effectName; bv = b.effectName; }
            else                    { av = isEnabled(a) ? 1 : 0; bv = isEnabled(b) ? 1 : 0; }
            if (typeof av === "string") { av = av.toLowerCase(); bv = bv.toLowerCase(); }
            // Primary sort, with stable secondary tie-break to keep
            // same-value rows visually grouped (comp, then layer, then effect).
            if (av !== bv) return (av < bv ? -1 : 1) * sortDir;
            if (a.compName  !== b.compName)  return a.compName  < b.compName  ? -1 : 1;
            if (a.layerName !== b.layerName) return a.layerName < b.layerName ? -1 : 1;
            return a.effectName < b.effectName ? -1 : (a.effectName > b.effectName ? 1 : 0);
        });
    }
    sortEntries();

    // ── UI ───────────────────────────────────────────────────────────────────
    var dlg = new Window("dialog", "Color-Modification Effects");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 10;
    dlg.margins = 14;

    dlg.add("statictext", undefined,
        entries.length + " color effect" + (entries.length === 1 ? "" : "s")
        + " in " + scanLabel + " (recursive). Select rows and toggle, or use the bulk buttons.");

    // Sort button row — ScriptUI listbox headers aren't clickable, so we
    // expose sort as a button row mimicking that interaction. Click a
    // column to sort by it; click the same column again to flip direction.
    var sortRow = dlg.add("group");
    sortRow.orientation = "row"; sortRow.alignChildren = ["left", "center"];
    sortRow.spacing = 4; sortRow.margins = [0, 0, 0, 0];
    sortRow.add("statictext", undefined, "Sort:");
    var SORT_LABELS = ["Comp", "Layer", "Effect", "State"];
    var sortBtns = [];
    for (var sb = 0; sb < SORT_LABELS.length; sb++) {
        var b = sortRow.add("button", undefined, SORT_LABELS[sb]);
        b.preferredSize = [95, 22];
        b.__col = sb;
        b.onClick = (function (col) {
            return function () {
                if (sortKey === col) sortDir = -sortDir; else { sortKey = col; sortDir = 1; }
                sortEntries();
                repopulate();
                refreshSortBtnLabels();
            };
        })(sb);
        sortBtns.push(b);
    }
    function refreshSortBtnLabels() {
        for (var i = 0; i < sortBtns.length; i++) {
            var arrow = (i === sortKey) ? (sortDir === 1 ? "  ↓" : "  ↑") : "";
            sortBtns[i].text = SORT_LABELS[i] + arrow;
        }
    }
    refreshSortBtnLabels();

    var lb = dlg.add("listbox", undefined, undefined, {
        multiselect: true,
        numberOfColumns: 4,
        showHeaders: true,
        columnTitles: ["Comp", "Layer", "Effect", "State"],
        columnWidths: [180, 180, 180, 70]
    });
    lb.preferredSize = [640, 560];

    function populateRow(idx) {
        var ent  = entries[idx];
        var item = lb.add("item", ent.compName);
        item.subItems[0].text = ent.layerName;
        item.subItems[1].text = ent.effectName;
        item.subItems[2].text = stateText(isEnabled(ent));
    }
    function repopulate() {
        // Clearing + re-adding loses selection by design — after a re-sort
        // the indices don't mean the same rows anyway.
        try { lb.removeAll(); } catch (eRA) {}
        for (var r2 = 0; r2 < entries.length; r2++) populateRow(r2);
    }
    for (var r = 0; r < entries.length; r++) populateRow(r);

    function refreshRow(idx) {
        try {
            lb.items[idx].subItems[2].text = stateText(entries[idx].effect.enabled);
        } catch (eRef) {}
    }

    function setEnabled(ent, idx, on) {
        try { ent.effect.enabled = on; } catch (eSE) {}
        refreshRow(idx);
    }

    function getSelectedIndices() {
        var out = [];
        var sel = lb.selection;
        if (!sel) return out;
        if (!(sel instanceof Array)) sel = [sel];
        for (var i = 0; i < sel.length; i++) out.push(sel[i].index);
        return out;
    }

    function forEachIndex(indices, fn) {
        if (!indices || indices.length === 0) return;
        app.beginUndoGroup("Toggle Color Effects");
        try {
            for (var i = 0; i < indices.length; i++) {
                var idx = indices[i];
                fn(entries[idx], idx);
            }
        } finally { app.endUndoGroup(); }
        // ScriptUI quirk: subItem.text changes on a SELECTED row don't
        // repaint until the row loses focus. Nudge the selection to force
        // an immediate redraw without changing which rows are selected.
        var selIdx = getSelectedIndices();
        try { lb.selection = null; } catch (eSR) {}
        if (selIdx.length > 0) {
            try {
                var reSel = [];
                for (var q = 0; q < selIdx.length; q++) reSel.push(selIdx[q]);
                lb.selection = reSel;
            } catch (eReSel) {}
        }
    }

    // ── Buttons ──────────────────────────────────────────────────────────────
    var btnGrp = dlg.add("group");
    btnGrp.orientation = "row"; btnGrp.alignment = ["fill", "bottom"];
    btnGrp.spacing = 8; btnGrp.margins = [0, 4, 0, 0];

    var btnToggleSel  = btnGrp.add("button", undefined, "Toggle Selected");   btnToggleSel.preferredSize  = [130, 28];
    var btnEnableSel  = btnGrp.add("button", undefined, "Enable Selected");   btnEnableSel.preferredSize  = [130, 28];
    var btnDisableSel = btnGrp.add("button", undefined, "Disable Selected");  btnDisableSel.preferredSize = [130, 28];
    btnGrp.add("statictext", undefined, "").alignment = ["fill", "center"];
    var btnEnableAll  = btnGrp.add("button", undefined, "Enable All");        btnEnableAll.preferredSize  = [100, 28];
    var btnDisableAll = btnGrp.add("button", undefined, "Disable All");       btnDisableAll.preferredSize = [100, 28];
    var btnClose      = btnGrp.add("button", undefined, "Close");             btnClose.preferredSize      = [80,  28];

    btnToggleSel.onClick = function () {
        var idx = getSelectedIndices();
        if (idx.length === 0) return;
        forEachIndex(idx, function (ent, i) {
            var cur = true;
            try { cur = ent.effect.enabled; } catch (eC) {}
            setEnabled(ent, i, !cur);
        });
    };
    btnEnableSel.onClick = function () {
        forEachIndex(getSelectedIndices(), function (ent, i) { setEnabled(ent, i, true); });
    };
    btnDisableSel.onClick = function () {
        forEachIndex(getSelectedIndices(), function (ent, i) { setEnabled(ent, i, false); });
    };
    btnEnableAll.onClick = function () {
        var all = [];
        for (var i = 0; i < entries.length; i++) all.push(i);
        forEachIndex(all, function (ent, i) { setEnabled(ent, i, true); });
    };
    btnDisableAll.onClick = function () {
        var all = [];
        for (var i = 0; i < entries.length; i++) all.push(i);
        forEachIndex(all, function (ent, i) { setEnabled(ent, i, false); });
    };
    btnClose.onClick = function () { dlg.close(); };

    dlg.show();

})();
