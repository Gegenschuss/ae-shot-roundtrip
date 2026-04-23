/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Language Preflight
 *
 * Audit companion to Toggle Language + Render Languages. Scans every
 * comp in the project for language-tagged layers and shows them in a
 * multi-column list so you can confirm that every comp/layer is
 * correctly tagged BEFORE you run a batch render.
 *
 * Detection is identical to Toggle Language — layer.comment starting
 * with "lang:XX" or layer name ending in "_lang_XX", case-insensitive.
 *
 * The dialog shows one row per tagged layer with:
 *   • Comp name (containing composition)
 *   • Layer name
 *   • Language code (lowercased/normalised)
 *   • Source of the tag (comment / name)
 *   • Current state (on / off)
 *
 * "Reveal selected" opens the selected layer's comp and selects the
 * layer inside it — useful for jumping to a suspect tag. "Copy as
 * text" dumps the whole list to the clipboard so you can paste the
 * audit into a note/ticket.
 *
 * Read-only: the script never mutates the project.
 */

(function () {

    var proj = app.project;

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
                if (m) return { code: m[1].toLowerCase(), source: "comment" };
            }
        } catch (eC) {}
        var nm = "";
        try { nm = String(layer.name || ""); } catch (eN) {}
        var m2 = NAME_RE.exec(nm);
        return m2 ? { code: m2[1].toLowerCase(), source: "name" } : null;
    }

    // ── scan ───────────────────────────────────────────────────────────

    var entries = [];     // { comp, layer, code, source }
    var byCode  = {};     // code → count
    var compSet = {};     // comp.id → true, to count unique comps
    var bySource = { comment: 0, name: 0 };

    for (var i = 1; i <= proj.numItems; i++) {
        var it = proj.item(i);
        if (!(it instanceof CompItem)) continue;
        for (var li = 1; li <= it.numLayers; li++) {
            var L = it.layer(li);
            var tag = detectLang(L);
            if (!tag) continue;
            entries.push({ comp: it, layer: L, code: tag.code, source: tag.source });
            byCode[tag.code] = (byCode[tag.code] || 0) + 1;
            compSet[it.id] = true;
            bySource[tag.source]++;
        }
    }

    if (entries.length === 0) {
        greyAlert("Language Preflight",
              "No language-tagged layers found in this project.\n\n"
            + "Tag layers with either:\n"
            + "    layer.comment starting with \"lang:XX\", or\n"
            + "    layer name ending in \"_lang_XX\"");
        return;
    }

    // Sort state. Columns: 0 = Comp, 1 = Layer, 2 = Lang, 3 = Via, 4 = State.
    // Default: by comp name, ascending. Secondary tie-break always follows
    // (comp → layer → code) so same-value rows stay visually grouped.
    var sortKey = 0;
    var sortDir = 1;  // 1 = asc, -1 = desc
    function sortEntries() {
        entries.sort(function (a, b) {
            var av, bv;
            if (sortKey === 0)      { av = a.comp.name;         bv = b.comp.name;         }
            else if (sortKey === 1) { av = String(a.layer.name); bv = String(b.layer.name); }
            else if (sortKey === 2) { av = a.code;              bv = b.code;              }
            else if (sortKey === 3) { av = a.source;            bv = b.source;            }
            else {
                var ae = false, be = false;
                try { ae = !!a.layer.enabled; } catch (eA) {}
                try { be = !!b.layer.enabled; } catch (eB) {}
                av = ae ? 1 : 0; bv = be ? 1 : 0;
            }
            if (typeof av === "string") { av = av.toLowerCase(); bv = bv.toLowerCase(); }
            if (av !== bv) return (av < bv ? -1 : 1) * sortDir;
            if (a.comp.name !== b.comp.name) return a.comp.name < b.comp.name ? -1 : 1;
            if (a.code !== b.code) return a.code < b.code ? -1 : 1;
            var al = String(a.layer.name), bl = String(b.layer.name);
            return al < bl ? -1 : (al > bl ? 1 : 0);
        });
    }
    sortEntries();

    var codesList = [];
    for (var k in byCode) codesList.push(k);
    codesList.sort();

    // ── dialog ─────────────────────────────────────────────────────────

    var dlg = new Window("dialog", "Language Preflight");
    dlg.orientation = "column"; dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 10; dlg.margins = 14;

    // Summary panel
    var sp = dlg.add("panel", undefined, "Summary");
    sp.orientation = "column"; sp.alignChildren = ["fill", "top"];
    sp.margins = [12, 12, 12, 12]; sp.spacing = 4;
    var compCount = 0;
    for (var _ in compSet) compCount++;
    sp.add("statictext", undefined,
          "Total: " + entries.length + " tagged layer(s) across " + compCount + " comp(s).");
    var codeParts = [];
    for (var ci = 0; ci < codesList.length; ci++) {
        codeParts.push(codesList[ci] + " (" + byCode[codesList[ci]] + ")");
    }
    sp.add("statictext", undefined, "Codes: " + codeParts.join(", "));
    sp.add("statictext", undefined,
          "Detected via: comment = " + bySource.comment
        + ", name suffix = " + bySource.name);

    // Filter row
    var fRow = dlg.add("group");
    fRow.orientation = "row"; fRow.alignChildren = ["left", "center"]; fRow.spacing = 6;
    fRow.add("statictext", undefined, "Filter language:");
    var filterDD = fRow.add("dropdownlist", undefined, ["(all)"].concat(codesList));
    filterDD.selection = 0;
    fRow.add("statictext", undefined, "    Sort:");
    var SORT_LABELS = ["Comp", "Layer", "Lang", "Via", "State"];
    var sortBtns = [];
    for (var sb = 0; sb < SORT_LABELS.length; sb++) {
        var b = fRow.add("button", undefined, SORT_LABELS[sb]);
        b.preferredSize = [80, 22];
        b.onClick = (function (col) {
            return function () {
                if (sortKey === col) sortDir = -sortDir; else { sortKey = col; sortDir = 1; }
                sortEntries();
                populate(currentFilter);
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

    // Listbox with columns. Height scales with entry count, capped at
    // roughly the screen height minus room for the dialog's other panels
    // + the OS menu/taskbar. Width is generous enough to hold long comp
    // and layer names without horizontal scrolling.
    var lb = dlg.add("listbox", undefined, [], {
        numberOfColumns: 5,
        showHeaders: true,
        columnTitles: ["Comp", "Layer", "Lang", "Via", "State"],
        columnWidths: [300, 460, 70, 90, 80]
    });
    var screenH = 900;
    try { screenH = $.screens[0].bottom - $.screens[0].top; } catch (eSc) {}
    var maxListH = Math.max(320, screenH - 300);
    lb.preferredSize = [1120, Math.min(Math.max(entries.length * 18 + 24, 360), maxListH)];

    function populate(filterCode) {
        lb.removeAll();
        for (var j = 0; j < entries.length; j++) {
            var e = entries[j];
            if (filterCode && e.code !== filterCode) continue;
            var row = lb.add("item", e.comp.name);
            row.subItems[0].text = e.layer.name;
            row.subItems[1].text = e.code;
            row.subItems[2].text = e.source;
            var enabled = false;
            try { enabled = !!e.layer.enabled; } catch (eE) {}
            row.subItems[3].text = enabled ? "on" : "off";
            row.__entry = e;
        }
        // ScriptUI subItem refresh quirk — nudge selection so the columns
        // render after removeAll()+add().
        if (lb.items.length > 0) {
            var wasSel = lb.selection;
            lb.selection = null;
            if (wasSel) lb.selection = 0;
        }
    }
    var currentFilter = null;
    filterDD.onChange = function () {
        var pick = filterDD.selection ? filterDD.selection.text : "(all)";
        currentFilter = (pick === "(all)") ? null : pick;
        populate(currentFilter);
    };
    populate(currentFilter);

    // Buttons
    var btnGrp = dlg.add("group");
    btnGrp.orientation = "row"; btnGrp.alignment = ["fill", "bottom"]; btnGrp.spacing = 6;
    var gotoBtn = btnGrp.add("button", undefined, "Go to Comp");
    gotoBtn.preferredSize = [110, 26];
    var revealBtn = btnGrp.add("button", undefined, "Reveal Layer");
    revealBtn.preferredSize = [120, 26];
    var copyBtn = btnGrp.add("button", undefined, "Copy as text");
    copyBtn.preferredSize = [110, 26];
    btnGrp.add("statictext", undefined, "").alignment = ["fill", "center"];
    var closeBtn = btnGrp.add("button", undefined, "Close", { name: "ok" });
    closeBtn.preferredSize = [90, 26];

    gotoBtn.onClick = function () {
        var sel = lb.selection;
        if (!sel) { greyAlert("Language Preflight", "Select a row first."); return; }
        var e = sel.__entry;
        // Open the comp in the viewer without touching layer selection — just
        // jump there so you can inspect in context.
        try { e.comp.openInViewer(); } catch (eOV) {}
    };

    revealBtn.onClick = function () {
        var sel = lb.selection;
        if (!sel) { greyAlert("Language Preflight", "Select a row first."); return; }
        var e = sel.__entry;
        try { e.comp.openInViewer(); } catch (eOV) {}
        // Deselect all other layers first so the revealed one is the only one
        // selected — easy to spot in the timeline.
        try {
            for (var lx = 1; lx <= e.comp.numLayers; lx++) e.comp.layer(lx).selected = false;
            e.layer.selected = true;
        } catch (eSel) {}
    };

    copyBtn.onClick = function () {
        var lines = ["Comp\tLayer\tLang\tVia\tState"];
        var pick = filterDD.selection ? filterDD.selection.text : "(all)";
        var filt = (pick === "(all)") ? null : pick;
        for (var j = 0; j < entries.length; j++) {
            var e = entries[j];
            if (filt && e.code !== filt) continue;
            var enabled = false;
            try { enabled = !!e.layer.enabled; } catch (eE) {}
            lines.push(e.comp.name + "\t" + e.layer.name + "\t" + e.code + "\t" + e.source + "\t" + (enabled ? "on" : "off"));
        }
        var txt = lines.join("\n");
        // ExtendScript has no direct clipboard API — shell out via pbcopy on
        // macOS, clip on Windows. Fall back to showing the text in a dialog
        // if neither is available.
        var copied = false;
        try {
            if ($.os.indexOf("Mac") >= 0) {
                var safe = txt.replace(/'/g, "'\\''");
                system.callSystem("printf '%s' '" + safe + "' | pbcopy");
                copied = true;
            } else if ($.os.indexOf("Windows") >= 0) {
                var safeW = txt.replace(/"/g, '""');
                system.callSystem('cmd /c echo "' + safeW + '" | clip');
                copied = true;
            }
        } catch (eCp) {}
        if (!copied) {
            greyAlert("Language Preflight",
                "Clipboard copy not available on this OS.\n\n" + txt);
        }
    };

    closeBtn.onClick = function () { dlg.close(1); };

    dlg.show();

})();
