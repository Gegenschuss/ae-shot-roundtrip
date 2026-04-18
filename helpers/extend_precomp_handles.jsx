/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Extend Precomp Handles
 *
 * Adds 50 frames of handle at both the in and out point of the
 * selected precomp layer:
 *   - Extends the precomp duration by 100 frames
 *   - Shifts all layers inside the precomp forward by 50 frames
 *   - Adjusts the layer's startTime in the parent comp so the
 *     visible content stays aligned
 *
 * Single undo step.
 */

(function () {

    var HANDLE_FRAMES = 50;

    var comp = app.project.activeItem;
    if (!(comp instanceof CompItem)) {
        alert("Extend Precomp Handles: please open a composition first.");
        return;
    }

    var selected = comp.selectedLayers;
    if (selected.length !== 1) {
        alert("Extend Precomp Handles: select exactly one precomp layer.");
        return;
    }

    var layer = selected[0];
    if (!(layer.source instanceof CompItem)) {
        alert("Extend Precomp Handles: the selected layer is not a precomp.");
        return;
    }

    var precomp = layer.source;
    var fps = precomp.frameRate;
    var handleSec = HANDLE_FRAMES / fps;

    app.beginUndoGroup("Extend Precomp Handles (\u00b1" + HANDLE_FRAMES + "f)");

    try {
        // 0. Unlock all locked layers inside the precomp (re-lock after)
        var lockedLayers = [];
        for (var i = 1; i <= precomp.numLayers; i++) {
            if (precomp.layer(i).locked) {
                lockedLayers.push(i);
                precomp.layer(i).locked = false;
            }
        }

        // 1. Extend precomp duration by 2× handle
        precomp.duration += handleSec * 2;

        // 2. Shift all layers inside the precomp forward by the handle amount
        for (var i = 1; i <= precomp.numLayers; i++) {
            try {
                precomp.layer(i).startTime += handleSec;
            } catch (layerErr) {
                writeLn("Warning: could not shift layer " + i + " (" + precomp.layer(i).name + "): " + layerErr.message);
            }
        }

        // 3. Shift precomp comp markers forward by the handle amount
        var markers = precomp.markerProperty;
        if (markers.numKeys > 0) {
            var markerData = [];
            for (var m = 1; m <= markers.numKeys; m++) {
                var v = markers.keyValue(m);
                var mv = new MarkerValue(v.comment);
                mv.chapter     = v.chapter;
                mv.url         = v.url;
                mv.frameTarget = v.frameTarget;
                mv.cuePointName = v.cuePointName;
                mv.duration    = v.duration;
                if (typeof v.label !== "undefined") mv.label = v.label;
                if (typeof v.protectedRegion !== "undefined") mv.protectedRegion = v.protectedRegion;
                markerData.push({ time: markers.keyTime(m), value: mv });
            }
            for (var m = markers.numKeys; m >= 1; m--) {
                markers.removeKey(m);
            }
            for (var m = 0; m < markerData.length; m++) {
                markers.setValueAtTime(markerData[m].time + handleSec, markerData[m].value);
            }
        }

        // 4. Pull the start timecode back so the content keeps its original TC
        precomp.displayStartTime -= handleSec;

        // 5. Extend-contract: extend the layer's trim outward first,
        //    then shift startTime, then contract the in point back
        var origIn  = layer.inPoint;
        var origOut = layer.outPoint;

        // Extend trim outward to make room
        layer.inPoint  = Math.max(0, origIn - handleSec);
        layer.outPoint = Math.min(comp.duration, origOut + handleSec);

        // Shift startTime so content aligns with the shifted precomp
        layer.startTime -= handleSec;

        // Contract in/out back to original positions
        layer.inPoint  = origIn;
        layer.outPoint = origOut;

        // 6. Re-lock previously locked layers
        for (var i = 0; i < lockedLayers.length; i++) {
            precomp.layer(lockedLayers[i]).locked = true;
        }
    } finally {
        app.endUndoGroup();
    }

    writeLn("Extend Precomp Handles: added \u00b1" + HANDLE_FRAMES + " frames to \"" + precomp.name + "\".");

})();
