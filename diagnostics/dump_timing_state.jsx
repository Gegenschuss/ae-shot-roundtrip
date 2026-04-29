/**
 * dump_timing_state.jsx
 *
 * Diagnostic snapshot of every AVLayer's timing state inside the active
 * comp AND every nested precomp.  Run BEFORE and AFTER a shot-roundtrip
 * pass; diff the two text files to see exactly what the roundtrip
 * mutated (and what it failed to mutate).
 *
 * Usage:
 *   1. Open the comp you want to dump.  Make it the active item.
 *   2. File > Scripts > Run Script File... > dump_timing_state.jsx
 *   3. Tag the snapshot ("before" / "after" / "post-bake" etc.).
 *   4. The dump lands next to the .aep as
 *      timing_dump_<compName>_<tag>.txt  (UTF-8, plain text).
 *   5. Repeat with a different tag after each step you want to inspect.
 *
 * What it captures per layer:
 *   - inPoint / outPoint / startTime / duration (4-decimal seconds)
 *   - timeStretch
 *   - timeRemapEnabled + numKeys + per-key (time, value, in/out interp)
 *   - remap value sampled at inPoint and outPoint (direction check)
 *   - source name + source duration
 *   - layer markers (cut in / cut out etc.)
 *   - parent linkage
 *   - flags: enabled, solo, guide, adjustment
 *
 * Plus comp-level: name, duration, frameRate, comp markers.
 *
 * Tags layers that have any time effect with [TIME-EFFECT] so they're
 * easy to grep / scroll to.
 *
 * No external dependencies.  Read-only — does not mutate the project.
 */

(function () {

    if (!app.project || !app.project.activeItem ||
        !(app.project.activeItem instanceof CompItem)) {
        alert("Open a comp first, then run this script.");
        return;
    }
    var topComp = app.project.activeItem;

    // ── Tag dialog ──────────────────────────────────────────────────────
    var w = new Window("dialog", "Dump timing state");
    w.alignChildren = ["fill", "top"];
    w.margins = 14; w.spacing = 8;
    w.add("statictext", undefined, "Tag for this snapshot (used in the filename):");
    var inp = w.add("edittext", undefined, "before");
    inp.preferredSize.width = 280;
    var grp = w.add("group");
    grp.alignment = "right";
    var btnCancel = grp.add("button", undefined, "Cancel");
    var btnOK     = grp.add("button", undefined, "Save dump");
    btnOK.preferredSize     = [110, 24];
    btnCancel.preferredSize = [90,  24];
    btnOK.onClick     = function () { w.close(1); };
    btnCancel.onClick = function () { w.close(2); };
    if (w.show() !== 1) return;

    var tag = (inp.text || "snapshot").replace(/[^a-zA-Z0-9_-]/g, "_");

    // ── Output buffer ───────────────────────────────────────────────────
    var lines = [];
    function L(s) { lines.push(s); }
    function pad(n, w) {
        var s = String(n);
        while (s.length < w) s = " " + s;
        return s;
    }
    function fmt(n) {
        if (typeof n !== "number" || isNaN(n)) return String(n);
        return n.toFixed(4);
    }
    function interpName(t) {
        try {
            if (t === KeyframeInterpolationType.LINEAR) return "linear";
            if (t === KeyframeInterpolationType.BEZIER) return "bezier";
            if (t === KeyframeInterpolationType.HOLD)   return "hold";
        } catch (e) {}
        return "?(" + t + ")";
    }
    function safeName(s) {
        return String(s || "").replace(/[^a-zA-Z0-9_]/g, "_").substring(0, 60);
    }

    L("# dump_timing_state.jsx");
    L("# tag:        " + tag);
    L("# date:       " + (new Date()).toString());
    L("# project:    " + (app.project.file ? app.project.file.fsName : "<unsaved>"));
    L("# top comp:   " + topComp.name);
    L("# AE version: " + app.version);
    L("");

    // ── Layer dump ──────────────────────────────────────────────────────
    function classifyLayer(lyr) {
        try { if (lyr instanceof CameraLayer) return "Camera"; } catch (e) {}
        try { if (lyr instanceof LightLayer)  return "Light";  } catch (e) {}
        try { if (lyr instanceof TextLayer)   return "Text";   } catch (e) {}
        try { if (lyr instanceof ShapeLayer)  return "Shape";  } catch (e) {}
        try { if (lyr.nullLayer)              return "Null";   } catch (e) {}
        try { if (lyr.adjustmentLayer)        return "Adjust"; } catch (e) {}
        try { if (lyr.source instanceof CompItem) return "Precomp"; } catch (e) {}
        try {
            if (lyr.source && lyr.source.mainSource) {
                if (lyr.source.mainSource instanceof SolidSource) return "Solid";
                if (lyr.source.mainSource instanceof FileSource)  return "Footage";
            }
        } catch (e) {}
        return "Layer";
    }

    function dumpLayer(lyr) {
        var typ = classifyLayer(lyr);
        var hasTimeEffect = false;
        try { hasTimeEffect = (Math.abs(lyr.stretch - 100) > 0.01) || lyr.timeRemapEnabled; }
        catch (e) {}
        var tag = hasTimeEffect ? "  [TIME-EFFECT]" : "";

        L("Layer " + pad(lyr.index, 3) + ": '" + lyr.name + "'  [" + typ + "]" + tag);
        try { L("  inPoint:        " + fmt(lyr.inPoint));   } catch (e) {}
        try { L("  outPoint:       " + fmt(lyr.outPoint));  } catch (e) {}
        try { L("  duration:       " + fmt(lyr.outPoint - lyr.inPoint)); } catch (e) {}
        try { L("  startTime:      " + fmt(lyr.startTime)); } catch (e) {}
        try { L("  enabled:        " + lyr.enabled);        } catch (e) {}
        try { L("  solo:           " + lyr.solo);           } catch (e) {}
        try { L("  guide:          " + lyr.guideLayer);     } catch (e) {}
        try { L("  adjustment:     " + lyr.adjustmentLayer);} catch (e) {}
        try { L("  stretch:        " + lyr.stretch);        } catch (e) {}
        try { L("  parent:         " + (lyr.parent ? lyr.parent.name : "<none>")); } catch (e) {}
        try {
            if (lyr.source) {
                var srcKind = (lyr.source instanceof CompItem) ? " (PRECOMP)" : "";
                L("  source:         '" + lyr.source.name + "'  duration: " + fmt(lyr.source.duration) + srcKind);
            } else {
                L("  source:         <none>");
            }
        } catch (e) {}

        var trEn = false;
        try { trEn = !!lyr.timeRemapEnabled; } catch (e) {}
        L("  timeRemapEnabled: " + trEn);
        if (trEn) {
            try {
                var tr = lyr.property("Time Remap");
                var n = tr.numKeys;
                L("  remapKeys:      " + n);
                for (var k = 1; k <= n; k++) {
                    var kt = 0, kv = 0, kI = -1, kO = -1;
                    try { kt = tr.keyTime(k);                          } catch (e) {}
                    try { kv = tr.keyValue(k);                         } catch (e) {}
                    try { kI = tr.keyInInterpolationType(k);           } catch (e) {}
                    try { kO = tr.keyOutInterpolationType(k);          } catch (e) {}
                    var tEase = "", oEase = "";
                    try {
                        var ti = tr.keyInTemporalEase(k);
                        if (ti && ti[0]) tEase = "  inEase[s=" + fmt(ti[0].speed) + ",i=" + fmt(ti[0].influence) + "]";
                    } catch (e) {}
                    try {
                        var to = tr.keyOutTemporalEase(k);
                        if (to && to[0]) oEase = "  outEase[s=" + fmt(to[0].speed) + ",i=" + fmt(to[0].influence) + "]";
                    } catch (e) {}
                    L("    " + pad(k, 2) + ": t=" + fmt(kt) +
                      "  v=" + fmt(kv) +
                      "  interp=" + interpName(kI) + "/" + interpName(kO) +
                      tEase + oEase);
                }
                // Sample at in/out for direction tag.
                var vIn  = 0, vOut = 0;
                try { vIn  = tr.valueAtTime(lyr.inPoint,  false); } catch (e) {}
                try { vOut = tr.valueAtTime(lyr.outPoint, false); } catch (e) {}
                var dir = "FLAT";
                if (vOut > vIn + 1e-6) dir = "ASCENDING";
                else if (vOut < vIn - 1e-6) dir = "DESCENDING";
                L("  remapAtIn:      " + fmt(vIn));
                L("  remapAtOut:     " + fmt(vOut));
                L("  direction:      " + dir);
            } catch (eR) {
                L("  (remap inspection error: " + eR + ")");
            }
        }

        // Layer markers — useful for cut-in / cut-out tagging.
        try {
            var lm = lyr.property("Marker");
            if (lm && lm.numKeys && lm.numKeys > 0) {
                L("  layerMarkers:   " + lm.numKeys);
                for (var mi = 1; mi <= lm.numKeys; mi++) {
                    var mt = 0, mv = null;
                    try { mt = lm.keyTime(mi);  } catch (e) {}
                    try { mv = lm.keyValue(mi); } catch (e) {}
                    var mc = (mv && mv.comment) ? mv.comment : "";
                    L("    @ " + fmt(mt) + "  '" + mc + "'");
                }
            }
        } catch (e) {}

        L("");
    }

    var visited = {};

    function dumpComp(comp, depth, breadcrumb) {
        var key = "c" + (comp.id || comp.name);
        if (visited[key]) {
            L("(comp '" + comp.name + "' already dumped above; skipping recursion)");
            L("");
            return;
        }
        visited[key] = true;

        L("=================================================================");
        L("=== Comp: " + comp.name + "  (depth " + depth + ")");
        if (breadcrumb && breadcrumb.length) L("=== via:  " + breadcrumb.join(" > "));
        L("=================================================================");
        try { L("  duration:       " + fmt(comp.duration));  } catch (e) {}
        try { L("  frameRate:      " + comp.frameRate);      } catch (e) {}
        try { L("  frameDuration:  " + fmt(1 / comp.frameRate)); } catch (e) {}
        try { L("  size:           " + comp.width + " x " + comp.height); } catch (e) {}
        try { L("  numLayers:      " + comp.numLayers);      } catch (e) {}

        try {
            if (comp.markerProperty && comp.markerProperty.numKeys > 0) {
                L("  compMarkers:    " + comp.markerProperty.numKeys);
                for (var mi = 1; mi <= comp.markerProperty.numKeys; mi++) {
                    var mt = comp.markerProperty.keyTime(mi);
                    var mv = comp.markerProperty.keyValue(mi);
                    var mc = (mv && mv.comment) ? mv.comment : "";
                    L("    @ " + fmt(mt) + "  '" + mc + "'");
                }
            }
        } catch (e) {}

        L("");

        var precomps = [];
        for (var i = 1; i <= comp.numLayers; i++) {
            var lyr;
            try { lyr = comp.layer(i); } catch (e) { continue; }
            if (!lyr) continue;
            dumpLayer(lyr);
            try {
                if (lyr.source instanceof CompItem) {
                    precomps.push({ comp: lyr.source, layerName: lyr.name });
                }
            } catch (e) {}
        }

        for (var pi = 0; pi < precomps.length; pi++) {
            dumpComp(precomps[pi].comp, depth + 1,
                     (breadcrumb || []).concat([comp.name + " > " + precomps[pi].layerName]));
        }
    }

    dumpComp(topComp, 0, []);

    // ── Save ────────────────────────────────────────────────────────────
    var saveDir;
    if (app.project.file) saveDir = app.project.file.parent;
    else                  saveDir = Folder.desktop;
    var fname = "timing_dump_" + safeName(topComp.name) + "_" + tag + ".txt";
    var f = new File(saveDir.fsName + "/" + fname);
    f.encoding = "UTF-8";
    f.open("w");
    f.write(lines.join("\n"));
    f.close();

    alert("Wrote " + lines.length + " lines:\n\n" + f.fsName +
          "\n\nRun again with a different tag (e.g. 'after') after the next step.");
})();
