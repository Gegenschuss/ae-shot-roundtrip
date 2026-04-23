/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Render Languages
 *
 * Batch-render one or more template render queue items across every
 * language tagged in the project. Uses the same tagging convention as
 * Toggle Language — layer.comment starting with "lang:XX" or layer
 * name ending in "_lang_XX" — and is CASE-INSENSITIVE (so "_LANG_DE",
 * "_Lang_de", "lang:DE" all work and normalize to "de"). Pick multiple
 * templates in the dialog (Cmd/Ctrl or Shift-click, or "Select all
 * QUEUED items" shortcut) and every template is rendered at every
 * picked language.
 *
 * Two render targets, picked via a radio in the dialog:
 *
 *   • AE Render Queue (default) — blocking, in-place. For each
 *     (template × language) pair: toggle → duplicate template → suffix
 *     output filename with "_<lang>" → disable other queue items →
 *     render (blocks until done) → next. Template-first order so a
 *     failure mid-batch leaves earlier comps with complete sets. Each
 *     language leaves a DONE queue item behind as an audit record.
 *     After the batch the project language is restored.
 *
 *   • Adobe Media Encoder — dispatch via Dynamic Link. AME reads the
 *     saved .aep file at render time, so one physical file per language
 *     is the only way to avoid cross-contamination. For each language:
 *     toggle the project → Save As "{project}_<lang>.aep" (variant file
 *     alongside the original) → queueInAME(false). The variant file
 *     holds all picked templates with language-suffixed output paths,
 *     and AME's queue holds N items (one per language) each pointing at
 *     its own immutable .aep. After dispatch the original .aep is
 *     reopened from disk so AE is back where you started. Optional
 *     "Start AME queue immediately" checkbox passes `true` on the final
 *     queueInAME call so AME kicks off without waiting for manual start.
 *
 * Skip-existing semantics: AE mode checks the output file on disk,
 * AME mode checks whether the variant .aep already exists (treated as
 * "already dispatched previously").
 */

(function () {

    var proj = app.project;
    var rq   = proj.renderQueue;

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

    // Palette progress window helper. A non-modal "palette" stays on
    // screen while the main script does its work. Between iterations we
    // call win.update() — that's when the Cancel button's onClick gets
    // a chance to fire. During AE's blocking rq.render() the palette
    // freezes (single-threaded ExtendScript), but AE's own render panel
    // takes over so the user still has ESC to stop the in-progress item.
    function createProgressWin(title, total) {
        var win = null;
        try {
            win = new Window("palette", title, undefined, { resizeable: false, independent: true });
        } catch (eNew) { return null; }

        win.orientation = "column"; win.alignChildren = ["fill", "top"];
        win.spacing = 8; win.margins = 14;

        var label = win.add("statictext", undefined, "Preparing…", { multiline: true });
        label.preferredSize = [460, 34];

        var bar = win.add("progressbar", undefined, 0, Math.max(1, total));
        bar.preferredSize = [460, 14];
        bar.value = 0;

        var countLabel = win.add("statictext", undefined, "0 / " + total);
        countLabel.alignment = ["right", "center"];

        var btnRow = win.add("group");
        btnRow.orientation = "row"; btnRow.alignment = ["fill", "bottom"];
        btnRow.add("statictext", undefined, "").alignment = ["fill", "center"];
        var cancelBtn = btnRow.add("button", undefined, "Cancel");
        cancelBtn.preferredSize = [100, 24];

        var state = { cancelled: false };
        cancelBtn.onClick = function () {
            state.cancelled = true;
            try { cancelBtn.text = "Cancelling…"; } catch (eCT) {}
            try { cancelBtn.enabled = false; } catch (eCE) {}
            try { win.update(); } catch (eCU) {}
        };

        try { win.center(); } catch (eC) {}
        try { win.show(); } catch (eS) {}
        try { win.update(); } catch (eU) {}

        return {
            setStatus: function (msg, doneCount) {
                try { label.text = msg; } catch (eL) {}
                if (typeof doneCount === "number") {
                    try { bar.value = doneCount; } catch (eB) {}
                    try { countLabel.text = doneCount + " / " + total; } catch (eCl) {}
                }
                try { win.update(); } catch (eU) {}
            },
            isCancelled: function () {
                try { win.update(); } catch (eU) {}
                return state.cancelled;
            },
            close: function () {
                try { win.close(); } catch (eC) {}
            }
        };
    }

    if (rq.numItems === 0) {
        greyAlert("Render Languages", "Add at least one item to the render queue first.");
        return;
    }

    // ── language detection (mirrors helpers/toggle_language.jsx) ──────

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

    // One scan, keyed by code → list of {comp, layer}. Used to apply
    // the toggle per language without re-walking the project each round.
    var tagged = {};
    for (var i = 1; i <= proj.numItems; i++) {
        var it = proj.item(i);
        if (!(it instanceof CompItem)) continue;
        for (var li = 1; li <= it.numLayers; li++) {
            var L = it.layer(li);
            var code = detectLang(L);
            if (!code) continue;
            if (!tagged[code]) tagged[code] = [];
            tagged[code].push({ comp: it, layer: L });
        }
    }

    var codes = [];
    for (var k in tagged) codes.push(k);
    codes.sort();

    if (codes.length === 0) {
        greyAlert("Render Languages",
              "No language-tagged layers found.\n\n"
            + "Tag layers via Toggle Language conventions first:\n"
            + "    layer.comment starting with \"lang:XX\", or\n"
            + "    layer name ending in \"_lang_XX\"");
        return;
    }

    function applyLanguage(activeCode) {
        for (var key in tagged) {
            var shouldEnable = (key === activeCode);
            var entries = tagged[key];
            for (var e = 0; e < entries.length; e++) {
                var layer = entries[e].layer;
                try { layer.enabled = shouldEnable; } catch (eE) {}
                try { if (layer.hasAudio) layer.audioEnabled = shouldEnable; } catch (eA) {}
            }
        }
    }

    // Best-effort detection of the current project language — whichever
    // code has any enabled tagged layer wins. Used to restore state
    // after the batch. Falls back to the first code if nothing is
    // currently enabled (e.g. project was set to an untagged language).
    function detectCurrentLanguage() {
        for (var c = 0; c < codes.length; c++) {
            var entries = tagged[codes[c]];
            for (var e = 0; e < entries.length; e++) {
                try { if (entries[e].layer.enabled) return codes[c]; } catch (eX) {}
            }
        }
        return codes[0];
    }

    var originalLang = detectCurrentLanguage();

    // ── filename helper ──────────────────────────────────────────────
    // Insert "_<LANG>" (UPPERCASE) before the last dot. Sequence patterns
    // like "Frames_[####].png" become "Frames_[####]_DE.png"; plain files
    // like "MyComp.mov" become "MyComp_DE.mov"; extension-less files get
    // the suffix appended. Internal lang codes stay lowercase for matching
    // across tag conventions; only the file-/.aep-name output is UPPERCASE
    // so deliverables read cleanly (industry convention).
    function withLangSuffix(fsName, lang) {
        var langUp = String(lang || "").toUpperCase();
        var slash = Math.max(fsName.lastIndexOf("/"), fsName.lastIndexOf("\\"));
        var dir  = slash >= 0 ? fsName.substring(0, slash + 1) : "";
        var base = slash >= 0 ? fsName.substring(slash + 1)    : fsName;
        var dot  = base.lastIndexOf(".");
        if (dot < 0) return dir + base + "_" + langUp;
        return dir + base.substring(0, dot) + "_" + langUp + base.substring(dot);
    }

    // ── dialog ────────────────────────────────────────────────────────

    var dlg = new Window("dialog", "Render Languages");
    dlg.orientation = "column"; dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 10; dlg.margins = 14;

    // Render target — picks the dispatch pipeline. AE mode (default) is
    // the original blocking in-place render. AME mode saves a
    // {project}_<lang>.aep per language and queues each in Adobe Media
    // Encoder via Dynamic Link — necessary because AME reads the saved
    // AEP file at render time, so one physical file per language state
    // is the only way to avoid cross-contamination.
    var targetPnl = dlg.add("panel", undefined, "Render target");
    targetPnl.orientation = "column"; targetPnl.alignChildren = ["left", "top"];
    targetPnl.margins = [10, 15, 10, 10]; targetPnl.spacing = 4;
    var rbAE  = targetPnl.add("radiobutton", undefined,
        "AE Render Queue  —  blocking, renders in place (default)");
    var rbAME = targetPnl.add("radiobutton", undefined,
        "Adobe Media Encoder  —  saves {project}_<lang>.aep per language and queues each in AME");
    rbAE.value = true;
    var startAMECB = targetPnl.add("checkbox", undefined,
        "Start AME queue immediately after dispatch (AME mode only)");
    startAMECB.value = false;
    startAMECB.enabled = false;
    function refreshTargetEnabled() {
        startAMECB.enabled = rbAME.value;
    }
    rbAE.onClick  = refreshTargetEnabled;
    rbAME.onClick = refreshTargetEnabled;

    // Template list (multi-select)
    var tPnl = dlg.add("panel", undefined, "Template queue item(s)");
    tPnl.orientation = "column"; tPnl.alignChildren = ["fill", "top"];
    tPnl.margins = [10, 15, 10, 10]; tPnl.spacing = 6;

    var tHelp = tPnl.add("statictext", undefined,
          "Existing queue item(s) whose settings get cloned once per language. "
        + "Each picked template is rendered at each picked language — settings, "
        + "output module (codec/format), and output path are reused; only the "
        + "filename is rewritten with a _<lang> suffix. Templates themselves are "
        + "not rendered; duplicates are created per language. "
        + "Cmd/Ctrl-click to pick multiple, Shift-click for ranges.",
        { multiline: true });
    tHelp.preferredSize = [460, 68];

    var itemLabels = [];
    for (var r = 1; r <= rq.numItems; r++) {
        var rqi = rq.item(r);
        var stat = "";
        try {
            if      (rqi.status === RQItemStatus.QUEUED)       stat = "queued";
            else if (rqi.status === RQItemStatus.RENDERING)    stat = "rendering";
            else if (rqi.status === RQItemStatus.DONE)         stat = "done";
            else if (rqi.status === RQItemStatus.UNQUEUED)     stat = "unqueued";
            else if (rqi.status === RQItemStatus.ERR_STOPPED)  stat = "err-stopped";
            else if (rqi.status === RQItemStatus.USER_STOPPED) stat = "user-stopped";
            else stat = String(rqi.status);
        } catch (eSt) {}
        itemLabels.push(r + ". " + rqi.comp.name + "  [" + stat + "]");
    }
    var tLB = tPnl.add("listbox", undefined, itemLabels, { multiselect: true });
    tLB.preferredSize = [460, Math.min(Math.max(itemLabels.length * 18 + 8, 90), 200)];

    // Prefer every QUEUED item with render=true as default pick.
    var anySelected = false;
    for (var r2 = 0; r2 < tLB.items.length; r2++) {
        var rqi2 = rq.item(r2 + 1);
        var queuedAndOn = false;
        try { queuedAndOn = (rqi2.status === RQItemStatus.QUEUED && rqi2.render); } catch (ePe) {}
        tLB.items[r2].selected = queuedAndOn;
        if (queuedAndOn) anySelected = true;
    }
    if (!anySelected && tLB.items.length > 0) tLB.items[0].selected = true;

    // "All queued" shortcut — sets the listbox selection to every queued item
    // in one click. User can then deselect/add further.
    var allQueuedCB = tPnl.add("checkbox", undefined, "Select all QUEUED items");
    allQueuedCB.onClick = function () {
        if (!allQueuedCB.value) return;
        for (var i = 0; i < tLB.items.length; i++) {
            var it2 = rq.item(i + 1);
            var q = false;
            try { q = (it2.status === RQItemStatus.QUEUED); } catch (eQ) {}
            tLB.items[i].selected = q;
        }
        allQueuedCB.value = false; // one-shot action
        refreshPathPreview();
    };

    // Output path preview (shows the first selected template as a sample).
    var pathLabel = tPnl.add("statictext", undefined, "", { multiline: true });
    pathLabel.preferredSize = [460, 48];
    function getSelectedIndices() {
        var out = [];
        for (var s = 0; s < tLB.items.length; s++) {
            if (tLB.items[s].selected) out.push(s);
        }
        return out;
    }
    function refreshPathPreview() {
        var idxs = getSelectedIndices();
        if (idxs.length === 0) { pathLabel.text = "(no templates selected)"; return; }
        var item = rq.item(idxs[0] + 1);
        var sample = null;
        try { sample = item.outputModule(1).file; } catch (eO) {}
        var extra = idxs.length > 1 ? "   (+ " + (idxs.length - 1) + " more template" + (idxs.length > 2 ? "s" : "") + ")" : "";
        if (sample) {
            pathLabel.text = "Sample template output: " + sample.fsName + extra
                + "\nWill write: " + withLangSuffix(sample.fsName, "<lang>");
        } else {
            pathLabel.text = "(output module has no file path set — render will probably fail)" + extra;
        }
    }
    tLB.onChange = refreshPathPreview;
    refreshPathPreview();

    // Language checkbox list
    var lPnl = dlg.add("panel", undefined, "Languages");
    lPnl.orientation = "column"; lPnl.alignChildren = ["fill", "top"];
    lPnl.margins = [10, 15, 10, 10]; lPnl.spacing = 4;
    var lHelp = lPnl.add("statictext", undefined,
          "Detected from layer tags (layer.comment \"lang:XX\" or layer name "
        + "suffix \"_lang_XX\", case-insensitive). One render is produced per "
        + "checked language, with the project toggled to that language before "
        + "each render. Run \"Language Preflight\" (Little Toolbox) first if "
        + "you want to see every tagged layer before committing to a batch.",
        { multiline: true });
    lHelp.preferredSize = [460, 60];
    var langChecks = [];
    for (var c = 0; c < codes.length; c++) {
        var cb = lPnl.add("checkbox", undefined,
            codes[c] + "  (" + tagged[codes[c]].length + " layer" + (tagged[codes[c]].length === 1 ? "" : "s") + ")");
        cb.value = true;
        cb.__code = codes[c];
        langChecks.push(cb);
    }

    // Options
    var oPnl = dlg.add("panel", undefined, "Options");
    oPnl.orientation = "column"; oPnl.alignChildren = ["fill", "top"];
    oPnl.margins = [10, 15, 10, 10]; oPnl.spacing = 4;
    var chkSkipExisting = oPnl.add("checkbox", undefined,
        "Skip a language if the target file already exists");
    chkSkipExisting.value = true;

    var note = dlg.add("statictext", undefined,
        "AE mode: each language leaves a DONE queue item behind as a record; "
        + "the project language is restored afterward. "
        + "AME mode: each language writes a {project}_<lang>.aep next to the "
        + "original .aep and queues it in AME; the original .aep is reopened "
        + "from disk at the end so you're back where you started.",
        { multiline: true });
    note.preferredSize = [460, 60];

    var btnGrp = dlg.add("group");
    btnGrp.orientation = "row"; btnGrp.alignment = ["fill", "bottom"];
    btnGrp.add("statictext", undefined, "").alignment = ["fill", "center"];
    var cancel = btnGrp.add("button", undefined, "Cancel", { name: "cancel" });
    cancel.preferredSize = [80, 28];
    var render = btnGrp.add("button", undefined, "Render", { name: "ok" });
    render.preferredSize = [110, 28];

    cancel.onClick = function () { dlg.close(2); };
    render.onClick = function () {
        if (getSelectedIndices().length === 0) {
            greyAlert("Render Languages", "Pick at least one template queue item.");
            return;
        }
        var any = false;
        for (var cb2 = 0; cb2 < langChecks.length; cb2++) {
            if (langChecks[cb2].value) { any = true; break; }
        }
        if (!any) { greyAlert("Render Languages", "Pick at least one language."); return; }
        dlg.close(1);
    };

    if (dlg.show() !== 1) return;

    var renderTarget  = rbAME.value ? "ame" : "ae";
    var startAMEAtEnd = startAMECB.value;

    var templates = [];
    var idxs = getSelectedIndices();
    for (var ti = 0; ti < idxs.length; ti++) {
        templates.push(rq.item(idxs[ti] + 1));
    }
    var pickedCodes = [];
    for (var pc = 0; pc < langChecks.length; pc++) {
        if (langChecks[pc].value) pickedCodes.push(langChecks[pc].__code);
    }
    var skipExisting = chkSkipExisting.value;

    // ── execute ──────────────────────────────────────────────────────

    // Per-template report: { compName, rendered: [lang], skipped: [msg], errors: [msg] }
    // Populated by the AE path; the AME path populates `ameReport` instead.
    var report = [];
    var totalRendered = 0, totalSkipped = 0, totalErrors = 0;
    var ameReport = null;   // set by runAMEDispatch when renderTarget === "ame"
    var userAborted = false;

    var totalWork = (renderTarget === "ae")
                  ? (templates.length * pickedCodes.length)
                  : pickedCodes.length;
    var progressTitle = (renderTarget === "ae")
                      ? "Render Languages — AE Render Queue"
                      : "Render Languages — Adobe Media Encoder";
    var progress = createProgressWin(progressTitle, totalWork);

    if (renderTarget === "ae") {

    // Snapshot + disable every queue item so only our duplicates render.
    // Save originals so we can restore after the batch.
    var renderFlagBackup = [];
    for (var q = 1; q <= rq.numItems; q++) {
        try { renderFlagBackup.push({ item: rq.item(q), was: rq.item(q).render }); }
        catch (eB) {}
    }
    for (var qb = 0; qb < renderFlagBackup.length; qb++) {
        try { renderFlagBackup[qb].item.render = false; } catch (eD) {}
    }

    // Outer loop = templates. Template-first order so a failure mid-batch
    // leaves earlier templates fully rendered (complete sets per comp).
    for (var tmpI = 0; tmpI < templates.length; tmpI++) {
        var template = templates[tmpI];
        var tName = template.comp ? template.comp.name : "(no comp)";
        var tReport = { compName: tName, rendered: [], skipped: [], errors: [] };
        report.push(tReport);

        // Snapshot this template's output file per output module — each
        // template can have its own output path convention.
        var numOM = template.numOutputModules;
        var origOutputPaths = [];
        for (var om = 1; om <= numOM; om++) {
            var orig = null;
            try { orig = template.outputModule(om).file; } catch (eOF) {}
            origOutputPaths.push(orig);
        }

        for (var pl = 0; pl < pickedCodes.length; pl++) {
            var lang = pickedCodes[pl];
            var doneCount = tmpI * pickedCodes.length + pl;

            // Mid-batch cancel check — the user clicked Cancel in the
            // progress palette between renders. (During rq.render() the
            // palette is frozen; AE's own render panel handles ESC.)
            if (progress && progress.isCancelled()) {
                userAborted = true;
                break;
            }
            if (progress) {
                progress.setStatus(
                    "Rendering " + tName + "  /  " + lang
                    + "   (item " + (doneCount + 1) + " of " + totalWork + ")",
                    doneCount);
            }

            var sampleFsName = origOutputPaths[0] ? origOutputPaths[0].fsName : null;
            if (skipExisting && sampleFsName) {
                var target = new File(withLangSuffix(sampleFsName, lang));
                if (target.exists) {
                    tReport.skipped.push(lang + " — " + target.name + " already exists");
                    totalSkipped++;
                    continue;
                }
            }

            applyLanguage(lang);

            var dup;
            try {
                dup = template.duplicate();
            } catch (eDup) {
                tReport.errors.push(lang + " — could not duplicate template: " + eDup);
                totalErrors++;
                continue;
            }

            var renameOK = true;
            for (var om2 = 1; om2 <= numOM && om2 <= dup.numOutputModules; om2++) {
                var origPath = origOutputPaths[om2 - 1];
                if (!origPath) continue;
                try {
                    dup.outputModule(om2).file = new File(withLangSuffix(origPath.fsName, lang));
                } catch (eRN) {
                    tReport.errors.push(lang + " — rename output module " + om2 + " failed: " + eRN);
                    totalErrors++;
                    renameOK = false;
                    break;
                }
            }
            if (!renameOK) {
                try { dup.remove(); } catch (eRm) {}
                continue;
            }

            dup.render = true;

            try {
                rq.render();                     // blocks until this dup renders
                // Check whether the user pressed ESC during AE's render
                // dialog — AE sets status to USER_STOPPED and returns from
                // render() normally. Treat it as an abort signal for the
                // whole batch instead of marching on to the next item.
                var stopped = false;
                try { stopped = (dup.status === RQItemStatus.USER_STOPPED); } catch (eSt) {}
                if (stopped) {
                    tReport.errors.push(lang + " — cancelled (ESC) during render");
                    totalErrors++;
                    userAborted = true;
                } else {
                    tReport.rendered.push(lang);
                    totalRendered++;
                }
            } catch (eR) {
                tReport.errors.push(lang + " — render failed: " + eR);
                totalErrors++;
            }

            // Leave dup as a DONE record. Disable its render flag so a
            // future rq.render() doesn't pick it up again, and so the
            // NEXT template's own iteration doesn't render it.
            try { dup.render = false; } catch (eDF) {}

            if (userAborted) break;   // break out of languages loop
        }
        if (userAborted) break;       // break out of templates loop
    }

    // Restore original language.
    applyLanguage(originalLang);

    // Restore original render flags on pre-existing items.
    for (var rb = 0; rb < renderFlagBackup.length; rb++) {
        try { renderFlagBackup[rb].item.render = renderFlagBackup[rb].was; } catch (eRB) {}
    }

    } else {
        // ── AME dispatch ─────────────────────────────────────────────
        // Each language gets its own {project}_<lang>.aep on disk with
        // the language state baked in, and each is queued separately
        // in AME. AME reads from the saved .aep at render time, so a
        // distinct file per language is required — without it, later
        // saves would overwrite the state earlier queue items depend on.
        ameReport = runAMEDispatch(progress);
    }

    if (progress) progress.close();

    // OS-level probe for a running Media Encoder process. Used to decide
    // whether queueInAME(false) will actually be received — AME has to be
    // up for a false-arg dispatch to land. queueInAME(true) launches AME
    // itself (and also starts its queue), which is why we can't use (true)
    // "just to launch" without also auto-starting the render.
    function isAMERunning() {
        try {
            if ($.os.indexOf("Mac") >= 0) {
                var out = system.callSystem("pgrep -if 'Adobe Media Encoder' 2>/dev/null");
                return !!(out && out.replace(/\s/g, "") !== "");
            } else if ($.os.indexOf("Windows") >= 0) {
                var out2 = system.callSystem('tasklist /FI "IMAGENAME eq Adobe Media Encoder.exe" /NH 2>NUL');
                return !!(out2 && out2.toLowerCase().indexOf("adobe media encoder") !== -1);
            }
        } catch (eR) {}
        return false; // couldn't detect → safer to assume not running
    }

    function runAMEDispatch(progress) {
        // Preconditions
        if (!proj.file) {
            greyAlert("Render Languages",
                  "Media Encoder mode needs the project saved to disk first,\n"
                + "so language variants can be written alongside the original.");
            return null;
        }
        var canQueue = false;
        try { canQueue = !!rq.canQueueInAME; } catch (eCQ) {}
        if (!canQueue) {
            greyAlert("Render Languages",
                  "Adobe Media Encoder isn't available to receive the queue.\n"
                + "Make sure AME is installed and that at least one queue item\n"
                + "has its Render flag enabled in AE's render queue.");
            return null;
        }

        // If the user wants "queue but don't start", AME has to already be
        // running — queueInAME(false) silently no-ops when AME is closed.
        // (queueInAME(true) would launch AME but also start its queue,
        // defeating the point of unchecking "Start immediately".)
        if (!startAMEAtEnd && !isAMERunning()) {
            greyAlert("Render Languages",
                  "Adobe Media Encoder isn't currently running.\n\n"
                + "Either open AME first, or check \"Start AME queue immediately\" "
                + "in the dialog — that option launches AME and starts the render.\n\n"
                + "Without one of these, dispatches to AME are silently dropped.");
            return null;
        }

        var originalFile = proj.file;
        var origDirFs    = originalFile.parent.fsName;
        var origName     = originalFile.name;                // e.g. "MyProject.aep"
        var origDot      = origName.lastIndexOf(".");
        var origStem     = origDot > 0 ? origName.substring(0, origDot) : origName;
        var origExt      = origDot > 0 ? origName.substring(origDot)   : ".aep";

        // Snapshot each template's output module paths so we can re-derive
        // language-suffixed paths from the originals each iteration
        // (not compound suffixes across iterations).
        var origPathsByTemplate = [];
        for (var tA = 0; tA < templates.length; tA++) {
            var perModulePaths = [];
            var nOM = templates[tA].numOutputModules;
            for (var oA = 1; oA <= nOM; oA++) {
                var p = null;
                try { p = templates[tA].outputModule(oA).file; } catch (eOP) {}
                perModulePaths.push(p);
            }
            origPathsByTemplate.push(perModulePaths);
        }

        var results = []; // { lang, variantPath, queued, skipped, error }

        for (var pLa = 0; pLa < pickedCodes.length; pLa++) {
            var langA = pickedCodes[pLa];
            var langUp = langA.toUpperCase();
            var variantFile = new File(origDirFs + "/" + origStem + "_" + langUp + origExt);
            var entry = { lang: langA, variantPath: variantFile.fsName, queued: false, skipped: false, error: null };

            // Progress + cancel check at the top of each iteration.
            if (progress && progress.isCancelled()) {
                entry.error = "cancelled by user";
                userAborted = true;
                results.push(entry);
                break;
            }
            if (progress) {
                progress.setStatus(
                    "Dispatching " + langA + " → AME"
                    + "   (" + (pLa + 1) + " of " + pickedCodes.length + ")",
                    pLa);
            }

            // Skip if the expected OUTPUT file already exists on disk AND
            // skipExisting is on. Same semantics as AE mode — skipExisting
            // is about "have I already produced this deliverable?", so we
            // check the actual output target (from template #1, module #1)
            // with the language suffix, not the .aep variant. Whether AME
            // would overwrite on render depends on its preset; we opt out
            // upstream rather than relying on AME's behaviour.
            var sampleOut = null;
            if (templates.length > 0 && origPathsByTemplate[0].length > 0) {
                sampleOut = origPathsByTemplate[0][0];
            }
            if (skipExisting && sampleOut) {
                var targetOut = new File(withLangSuffix(sampleOut.fsName, langA));
                if (targetOut.exists) {
                    entry.skipped = true;
                    entry.error = "output exists: " + targetOut.name;
                    results.push(entry);
                    continue;
                }
            }

            // Toggle to this language (mutates in-memory layer.enabled).
            applyLanguage(langA);

            // Disable EVERY existing render queue item — including dups we
            // created in earlier iterations (still present in-memory because
            // we haven't closed the project yet). queueInAME dispatches
            // whichever items have render=true; keeping only this iteration's
            // fresh dups enabled is how we avoid cross-contamination.
            for (var q2 = 1; q2 <= rq.numItems; q2++) {
                try { rq.item(q2).render = false; } catch (eDF2) {}
            }

            // Duplicate each picked template — AME marks items as queued
            // after a queueInAME call and refuses to re-dispatch them, so
            // we queue DUPLICATES, leaving the originals untouched for the
            // next iteration to duplicate again. Each dup gets a fresh
            // QUEUED status and is ours to modify.
            var iterationDups = [];
            var renameFailed = false;
            for (var tD = 0; tD < templates.length; tD++) {
                var dup;
                try { dup = templates[tD].duplicate(); }
                catch (eDup) {
                    entry.error = "duplicate template #" + (tD + 1) + " failed: " + eDup;
                    renameFailed = true;
                    break;
                }
                iterationDups.push(dup);

                // Apply lang-suffixed output paths to the duplicate, derived
                // from the ORIGINAL template's paths (not compounded).
                var paths = origPathsByTemplate[tD];
                for (var oR = 0; oR < paths.length && oR < dup.numOutputModules; oR++) {
                    var origP = paths[oR];
                    if (!origP) continue;
                    try {
                        dup.outputModule(oR + 1).file =
                            new File(withLangSuffix(origP.fsName, langA));
                    } catch (eRP) {
                        entry.error = "rename output module " + (oR + 1) + ": " + eRP;
                        renameFailed = true;
                        break;
                    }
                }
                if (renameFailed) break;

                // Ensure this specific dup is render=true (templates.render
                // defaults true after duplicate, but belt-and-suspenders).
                try { dup.render = true; } catch (eDR) {}
            }
            if (renameFailed) {
                // Roll back: remove the dups we created this iteration so
                // the variant .aep doesn't carry half-finished state.
                for (var iR = 0; iR < iterationDups.length; iR++) {
                    try { iterationDups[iR].remove(); } catch (eRm) {}
                }
                results.push(entry);
                continue;
            }

            // Save the current in-memory state to a new .aep on disk. This
            // BINDS the project to the variant file; the NEXT iteration's
            // save goes to a different variant file. The original .aep is
            // never touched after the pre-loop save.
            try {
                proj.save(variantFile);
            } catch (eSave) {
                entry.error = "save " + variantFile.name + " failed: " + eSave;
                results.push(entry);
                continue;
            }

            // Sanity check: our dup must be render=true and AME must be
            // ready to accept. If either isn't, queueInAME silently does
            // nothing — log it so the summary dialog explains why the
            // variant made it to disk but never reached AME.
            var anyQueueable = false;
            for (var qCk = 1; qCk <= rq.numItems; qCk++) {
                try { if (rq.item(qCk).render) { anyQueueable = true; break; } } catch (eQC) {}
            }
            var canNow = false;
            try { canNow = !!rq.canQueueInAME; } catch (eCN) {}

            if (!anyQueueable) {
                entry.error = "no render-enabled items in RQ at dispatch time "
                            + "(dup.render=true may have failed)";
                results.push(entry);
                continue;
            }
            if (!canNow) {
                entry.error = "canQueueInAME was false at dispatch time "
                            + "(AME unavailable or queue not queueable)";
                results.push(entry);
                continue;
            }

            // Pass the user's startAMEAtEnd preference on every call:
            //   true  → AME launches if needed AND starts processing
            //   false → items are added to AME's queue without starting
            //           (requires AME to already be running — we checked
            //           this upfront in runAMEDispatch, so by this point
            //           AME is reachable).
            try {
                rq.queueInAME(startAMEAtEnd);
                entry.queued = true;
            } catch (eQ) {
                entry.error = "queueInAME failed: " + eQ;
            }
            results.push(entry);
        }

        // Restore the user's original on-disk project state by reopening
        // from disk. Discard the in-memory variant binding. This leaves
        // originalFile untouched.
        try { proj.close(CloseOptions.DO_NOT_SAVE_CHANGES); } catch (eC) {}
        try { app.open(originalFile); } catch (eO) {}

        return {
            originalFile: originalFile.fsName,
            results:      results,
            startedAME:   startAMEAtEnd
        };
    }

    // ── summary ──────────────────────────────────────────────────────
    (function showSummary() {
        var dlg2 = new Window("dialog", "Render Languages — Done");
        dlg2.orientation = "column"; dlg2.alignChildren = ["fill", "top"];
        dlg2.spacing = 10; dlg2.margins = 14;

        if (renderTarget === "ame") {
            showAMESummary(dlg2);
        } else {
            showAESummary(dlg2);
        }

        var bg = dlg2.add("group");
        bg.orientation = "row"; bg.alignment = ["fill", "bottom"];
        bg.add("statictext", undefined, "").alignment = ["fill", "center"];
        var ok = bg.add("button", undefined, "OK", { name: "ok" });
        ok.preferredSize = [90, 28];
        ok.onClick = function () { dlg2.close(1); };
        dlg2.show();
    })();

    function showAESummary(dlg2) {
        var sp = dlg2.add("panel", undefined, "Summary");
        sp.orientation = "column"; sp.alignChildren = ["fill", "top"];
        sp.margins = [12, 12, 12, 12]; sp.spacing = 4;
        sp.add("statictext", undefined, "Target: AE Render Queue");
        sp.add("statictext", undefined,
              "Templates: " + report.length
            + "   |   Languages: " + pickedCodes.length
            + "   |   Renders attempted: " + (report.length * pickedCodes.length));
        sp.add("statictext", undefined,
              "Rendered: " + totalRendered
            + "   |   Skipped: " + totalSkipped
            + "   |   Errors: " + totalErrors);
        if (userAborted) {
            sp.add("statictext", undefined, "Batch was ABORTED by user (ESC) — remaining renders skipped.");
        }
        sp.add("statictext", undefined, "Project language restored to: " + originalLang);

        // Per-template breakdown. Each template gets one line summarising
        // what rendered / skipped / errored for it.
        var bp = dlg2.add("panel", undefined, "Per-template");
        bp.orientation = "column"; bp.alignChildren = ["fill", "top"];
        bp.margins = [12, 12, 12, 12]; bp.spacing = 2;
        var lines = "";
        for (var rI = 0; rI < report.length; rI++) {
            var tr = report[rI];
            lines += tr.compName + "\n";
            lines += "  rendered: " + (tr.rendered.length ? tr.rendered.join(", ") : "—") + "\n";
            if (tr.skipped.length) lines += "  skipped:  " + tr.skipped.length + "\n";
            if (tr.errors.length)  lines += "  errors:   " + tr.errors.length + "\n";
        }
        var maxH1 = $.screens[0].bottom - $.screens[0].top - 400;
        var h1 = Math.min(Math.max(lines.split("\n").length * 15 + 10, 80), maxH1);
        var ta1 = bp.add("edittext", undefined, lines,
            { multiline: true, readonly: true, scrollable: true });
        ta1.preferredSize = [520, h1];

        // Flatten all skipped/errored rows (with template prefix) for detail
        // panels, only shown if there's something to report.
        function flatten(kind) {
            var out = [];
            for (var i = 0; i < report.length; i++) {
                var tr2 = report[i];
                var arr = (kind === "skipped") ? tr2.skipped : tr2.errors;
                for (var j = 0; j < arr.length; j++) {
                    out.push(tr2.compName + "  ·  " + arr[j]);
                }
            }
            return out;
        }
        function addDetail(title, arr) {
            if (!arr.length) return;
            var p = dlg2.add("panel", undefined, title);
            p.orientation = "column"; p.alignChildren = ["fill", "top"];
            p.margins = [12, 12, 12, 12]; p.spacing = 4;
            var body = "";
            for (var n = 0; n < arr.length; n++) body += "  " + arr[n] + "\n";
            var maxH = $.screens[0].bottom - $.screens[0].top - 500;
            var h = Math.min(Math.max(arr.length * 15 + 20, 60), maxH);
            var ta = p.add("edittext", undefined, body,
                { multiline: true, readonly: true, scrollable: true });
            ta.preferredSize = [520, h];
        }
        addDetail("Skipped", flatten("skipped"));
        addDetail("Errors",  flatten("errors"));
    }

    function showAMESummary(dlg2) {
        var sp = dlg2.add("panel", undefined, "Summary");
        sp.orientation = "column"; sp.alignChildren = ["fill", "top"];
        sp.margins = [12, 12, 12, 12]; sp.spacing = 4;
        sp.add("statictext", undefined, "Target: Adobe Media Encoder");
        if (!ameReport) {
            sp.add("statictext", undefined, "(AME dispatch aborted — see earlier messages)");
            return;
        }
        var nQueued = 0, nSkipped = 0, nErr = 0;
        for (var rIA = 0; rIA < ameReport.results.length; rIA++) {
            var e = ameReport.results[rIA];
            if (e.queued)       nQueued++;
            else if (e.skipped) nSkipped++;
            if (e.error)        nErr++;
        }
        sp.add("statictext", undefined,
              "Languages: " + pickedCodes.length
            + "   |   Templates per variant: " + templates.length
            + "   |   Variants dispatched: " + nQueued);
        sp.add("statictext", undefined,
              "Queued: " + nQueued
            + "   |   Skipped: " + nSkipped
            + "   |   Errors: " + nErr);
        sp.add("statictext", undefined,
              (ameReport.startedAME
                ? "AME queue has been started (render_immediately=true on the last dispatch)."
                : "AME queue is paused — switch to Media Encoder and hit Start Queue."));
        sp.add("statictext", undefined, "Original project reopened: " + ameReport.originalFile);

        // Per-language breakdown.
        var bp = dlg2.add("panel", undefined, "Per-language");
        bp.orientation = "column"; bp.alignChildren = ["fill", "top"];
        bp.margins = [12, 12, 12, 12]; bp.spacing = 2;
        var lines = "";
        for (var rIL = 0; rIL < ameReport.results.length; rIL++) {
            var er = ameReport.results[rIL];
            var status = er.queued ? "queued"
                       : er.skipped ? "skipped (variant already exists)"
                       : er.error   ? ("error: " + er.error)
                       : "unknown";
            lines += er.lang + "  —  " + er.variantPath + "\n";
            lines += "  " + status + "\n";
        }
        var maxH2 = $.screens[0].bottom - $.screens[0].top - 400;
        var h2 = Math.min(Math.max(lines.split("\n").length * 15 + 10, 80), maxH2);
        var ta2 = bp.add("edittext", undefined, lines,
            { multiline: true, readonly: true, scrollable: true });
        ta2.preferredSize = [620, h2];
    }

})();
