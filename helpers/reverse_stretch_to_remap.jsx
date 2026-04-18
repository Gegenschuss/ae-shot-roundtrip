/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/*
================================================================================
  REVERSE STRETCH → REMAP (standalone helper)
  After Effects ExtendScript
================================================================================

For each selected AVLayer with negative stretch, rewrite the reversal as a
time-remap equivalent:

  - Set stretch back to 100
  - Enable time remap
  - Replace AE's auto-created keyframes with a pair placed at the HANDLE
    boundaries (compStart − handleSec, endKeyTime + handleSec). Values are
    linearly extrapolated from the stretch-derived slope so playback stays
    reversed across the full handle range, clamped to [0, srcDur] so we
    never ask for source frames that don't exist.

Identical logic to Shot Roundtrip's auto-conversion (including the handle
prompt) — exposed as a standalone helper so you can run it on one layer at
a time and verify the result before running a full roundtrip.

Idempotent on already-remap-reversed layers (they get skipped). Negative
stretch with AE's default "Hold in Place: Layer In-point" is what's assumed;
other Hold modes may shift the layer's visual extent after stretch reset.
================================================================================
*/

(function () {
    var proj = app.project;
    if (!proj) { alert("Reverse Stretch \u2192 Remap: no project open."); return; }

    var comp = proj.activeItem;
    if (!(comp instanceof CompItem)) {
        alert("Reverse Stretch \u2192 Remap: open a composition first.");
        return;
    }

    var sel = comp.selectedLayers;
    if (sel.length === 0) {
        alert("Reverse Stretch \u2192 Remap: select at least one layer.");
        return;
    }

    var handleFrames = promptHandleFrames(50);
    if (handleFrames === null) return; // user cancelled

    function convertOne(layer) {
        if (!(layer instanceof AVLayer)) return "skipped (not an AV layer)";
        if (layer.timeRemapEnabled)      return "skipped (time remap already on)";
        if (!(layer.stretch < 0))        return "skipped (not stretch-reversed, stretch=" + layer.stretch + ")";

        var srcDur = (layer.source && layer.source.duration) ? layer.source.duration : 0;
        if (srcDur <= 0) return "skipped (no source duration)";

        // 1. Capture current (reversed) state.
        //    For negatively-stretched layers AE swaps the reported in/out:
        //      inPoint  = comp-LATER  edge (source-earlier side)
        //      outPoint = comp-EARLIER edge (source-later side)
        //    Normalize into comp-timeline order.
        var startT   = layer.startTime;
        var stretch  = layer.stretch;
        var rawIn    = layer.inPoint;
        var rawOut   = layer.outPoint;
        var frameDur = layer.containingComp.frameDuration;

        // sourceTime(compTime) = (compTime - startTime) * (100 / stretch)
        // For stretch = -100 this reduces to (startTime - compTime).
        // Empirically, for reversed layers AE anchors startTime one frame past
        // the last-rendered source frame, so subtract one frameDur to get the
        // source time that's actually on screen.
        var srcAtRawIn  = (rawIn  - startT) * (100 / stretch) - frameDur;
        var srcAtRawOut = (rawOut - startT) * (100 / stretch) - frameDur;

        var compStart, compEnd, srcAtStart, srcAtEnd;
        if (rawIn <= rawOut) {
            compStart  = rawIn;        compEnd  = rawOut;
            srcAtStart = srcAtRawIn;   srcAtEnd = srcAtRawOut;
        } else {
            compStart  = rawOut;       compEnd  = rawIn;
            srcAtStart = srcAtRawOut;  srcAtEnd = srcAtRawIn;
        }

        if (srcAtStart < 0)      srcAtStart = 0;
        if (srcAtStart > srcDur) srcAtStart = srcDur;
        if (srcAtEnd   < 0)      srcAtEnd   = 0;
        if (srcAtEnd   > srcDur) srcAtEnd   = srcDur;

        // 2. Reset stretch to forward. AE may relocate the layer's extent.
        layer.stretch = 100;

        // 3. Anchor the layer so it covers the original comp range again.
        layer.startTime = compStart;
        layer.inPoint   = compStart;
        layer.outPoint  = compEnd;

        // 4. Enable time remap. Access via matchname + select the layer and
        //    property — otherwise AE marks the property "hidden" and
        //    setValueAtTime throws.
        layer.selected = true;
        layer.timeRemapEnabled = true;
        var tr = layer.property("ADBE Time Remapping");
        tr.selected = true;

        // Two keys placed at the handle boundaries (not the cut boundaries)
        // so AE never has to extrapolate past the last key when downstream
        // pipelines extend the layer beyond the cut (e.g. the roundtrip's
        // _dynamicLink wrapper). Slope comes from the stretch-derived source
        // times, so this works for any constant negative stretch magnitude
        // (-100 → slope -1, -50 → -2, -200 → -0.5, -124 → ≈ -0.806). Values
        // clamp to [0, srcDur]; when the source runs out, clamping freezes on
        // the source edge, the physically correct behaviour.
        //
        // AE's outPoint is exclusive (the first NON-rendered frame), so the
        // inner end reference lands one frame before compEnd.
        var endKeyTime = compEnd - frameDur;
        var keySpan    = endKeyTime - compStart;
        var handleSec  = handleFrames / layer.containingComp.frameRate;
        var slope      = (keySpan > 0) ? ((srcAtEnd - srcAtStart) / keySpan) : 0;
        var preTime    = compStart  - handleSec;
        var postTime   = endKeyTime + handleSec;
        var preVal     = srcAtStart - slope * handleSec;
        var postVal    = srcAtEnd   + slope * handleSec;
        if (preVal  < 0)      preVal  = 0;
        if (preVal  > srcDur) preVal  = srcDur;
        if (postVal < 0)      postVal = 0;
        if (postVal > srcDur) postVal = srcDur;

        // Write our keys FIRST. Removing all auto-keys before writing can
        // leave the property in an unusable state, so add ours, then drop
        // any stray auto-keys.
        tr.setValueAtTime(preTime,  preVal);
        tr.setValueAtTime(postTime, postVal);

        for (var k = tr.numKeys; k >= 1; k--) {
            var kt = tr.keyTime(k);
            if (Math.abs(kt - preTime)  > 0.0005 &&
                Math.abs(kt - postTime) > 0.0005) {
                tr.removeKey(k);
            }
        }

        // Force LINEAR interpolation on both keys. AE's default for a fresh
        // setValueAtTime on time remap is BEZIER, which eases between keys —
        // wrong for constant-speed reversed playback.
        for (var ki = 1; ki <= tr.numKeys; ki++) {
            try {
                tr.setInterpolationTypeAtKey(ki, KeyframeInterpolationType.LINEAR, KeyframeInterpolationType.LINEAR);
            } catch (eInterp) {}
        }

        return "remapped: pre[" + preTime.toFixed(3) + "=" + preVal.toFixed(3) + "] " +
               "post[" + postTime.toFixed(3) + "=" + postVal.toFixed(3) + "] " +
               "cut src[" + srcAtStart.toFixed(3) + "\u2192" + srcAtEnd.toFixed(3) + "]";
    }

    app.beginUndoGroup("Reverse Stretch \u2192 Remap");
    var results = [];
    try {
        for (var i = 0; i < sel.length; i++) {
            var layer = sel[i];
            var outcome;
            try { outcome = convertOne(layer); }
            catch (eLayer) { outcome = "error: " + eLayer.toString(); }
            results.push(layer.name + " \u2192 " + outcome);
        }
    } finally {
        app.endUndoGroup();
    }

    alert("Reverse Stretch \u2192 Remap\n\n" + results.join("\n"));

    function promptHandleFrames(defaultVal) {
        var dlg = new Window("dialog", "Reverse Stretch \u2192 Remap");
        dlg.orientation = "column"; dlg.alignChildren = ["fill", "top"];
        dlg.margins = 14; dlg.spacing = 10;

        dlg.add("statictext", undefined,
            "Handle length to encode in the remap curve.", { multiline: true });

        var row = dlg.add("group");
        row.orientation = "row"; row.alignChildren = ["left", "center"]; row.spacing = 8;
        row.add("statictext", undefined, "Handle frames:");
        var input = row.add("edittext", undefined, "" + defaultVal);
        input.characters = 6;
        input.active = true;

        var hint = dlg.add("statictext", undefined,
            "Match what you'll use in Shot Roundtrip. Keys are placed at \u00b1handle frames from the cut boundaries so playback stays reversed when the layer is later extended for handles.",
            { multiline: true });
        hint.preferredSize = [420, 48];

        var btnGrp = dlg.add("group");
        btnGrp.orientation = "row"; btnGrp.alignment = ["fill", "bottom"];
        btnGrp.add("statictext", undefined, "").alignment = ["fill", "center"];
        var btnCancel = btnGrp.add("button", undefined, "Cancel"); btnCancel.preferredSize = [80, 28];
        var btnOK     = btnGrp.add("button", undefined, "Convert"); btnOK.preferredSize     = [100, 28];

        var picked = null;
        btnCancel.onClick = function () { dlg.close(2); };
        btnOK.onClick = function () {
            var n = parseInt(input.text, 10);
            if (isNaN(n) || n < 0) { alert("Handle frames must be a non-negative integer."); return; }
            picked = n;
            dlg.close(1);
        };

        if (dlg.show() !== 1) return null;
        return picked;
    }
})();
