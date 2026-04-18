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
  - Replace AE's auto-created keyframes with a pair that reproduces the
    layer's current reversed playback:
       at layer.inPoint   → source time that was visible there pre-conversion
       at layer.outPoint  → source time that was visible there pre-conversion
    (swapped relative to the post-stretch-reset forward default, so playback
    stays reversed)

Identical logic to Shot Roundtrip's auto-conversion — exposed as a standalone
helper so you can run it on one layer at a time and verify the result before
running a full roundtrip.

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

        // AE's outPoint is exclusive (the first NON-rendered frame), so the
        // end key has to land one frame before compEnd to line up with the
        // last actually-rendered frame.
        var endKeyTime = compEnd - frameDur;

        // Write our two keys FIRST (reversed playback: srcAtStart > srcAtEnd).
        // Removing all auto-keys before writing can leave the property in an
        // unusable state, so add ours, then drop any stray auto-keys.
        tr.setValueAtTime(compStart,  srcAtStart);
        tr.setValueAtTime(endKeyTime, srcAtEnd);

        for (var k = tr.numKeys; k >= 1; k--) {
            var kt = tr.keyTime(k);
            if (Math.abs(kt - compStart)  > 0.0005 &&
                Math.abs(kt - endKeyTime) > 0.0005) {
                tr.removeKey(k);
            }
        }

        return "remapped: comp[" + compStart.toFixed(3) + "\u2192" + endKeyTime.toFixed(3) + "] " +
               "src["  + srcAtStart.toFixed(3) + "\u2192" + srcAtEnd.toFixed(3) + "]";
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
})();
