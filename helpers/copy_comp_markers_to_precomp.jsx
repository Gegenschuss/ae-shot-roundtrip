/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Copy Comp Markers to Selected Precomp
 *
 * Takes all composition markers from the active comp and copies them
 * into the comp source of the currently selected precomp layer.
 * Existing markers in the target precomp are left untouched.
 *
 * Single undo step.
 */

(function () {

    var comp = app.project.activeItem;
    if (!(comp instanceof CompItem)) {
        alert("Copy Comp Markers: please open a composition first.");
        return;
    }

    var selected = comp.selectedLayers;
    if (selected.length !== 1) {
        alert("Copy Comp Markers: select exactly one precomp layer.");
        return;
    }

    var layer = selected[0];
    if (!(layer.source instanceof CompItem)) {
        alert("Copy Comp Markers: the selected layer is not a precomp.");
        return;
    }

    var srcMarkers = comp.markerProperty;
    if (srcMarkers.numKeys === 0) {
        alert("Copy Comp Markers: the active comp has no markers.");
        return;
    }

    var targetComp = layer.source;
    var dstMarkers = targetComp.markerProperty;
    var copied = 0;

    app.beginUndoGroup("Copy Comp Markers to Precomp");

    try {
        for (var i = 1; i <= srcMarkers.numKeys; i++) {
            var compTime  = srcMarkers.keyTime(i);
            var localTime = compTime - layer.startTime;
            var value     = srcMarkers.keyValue(i);

            var mv = new MarkerValue(value.comment);
            mv.chapter     = value.chapter;
            mv.url         = value.url;
            mv.frameTarget = value.frameTarget;
            mv.cuePointName = value.cuePointName;
            mv.duration    = value.duration;
            if (typeof value.label !== "undefined") {
                mv.label = value.label;
            }
            if (typeof value.protectedRegion !== "undefined") {
                mv.protectedRegion = value.protectedRegion;
            }

            dstMarkers.setValueAtTime(localTime, mv);
            copied++;
        }
    } finally {
        app.endUndoGroup();
    }

    writeLn("Copy Comp Markers: " + copied + " marker(s) copied to \"" + targetComp.name + "\".");

})();
