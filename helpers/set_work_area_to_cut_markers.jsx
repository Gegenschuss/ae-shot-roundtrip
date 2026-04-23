/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Set Work Area to Cut Markers
 *
 * Finds the composition markers named "cut in" and "cut out" on the
 * active comp and sets the work area to that span. Shot Roundtrip and
 * Re-render Plates both place those markers automatically, so after a
 * roundtrip this makes RAM-preview and timeline focus snap to the
 * editorial cut (same range the tool now sets the work area to at
 * creation time — this helper is the one-click fix for comps created
 * before the cut-range work-area change, or for any comp that has the
 * markers but the work area has drifted).
 *
 * Takes the first marker of each name. Case-insensitive match.
 *
 * Single undo step.
 */

(function () {

    var comp = app.project.activeItem;
    if (!(comp instanceof CompItem)) {
        alert("Set Work Area to Cut Markers: open a composition first.");
        return;
    }

    var cutIn  = null;
    var cutOut = null;

    try {
        var mp = comp.markerProperty;
        if (mp && mp.numKeys > 0) {
            for (var i = 1; i <= mp.numKeys; i++) {
                var val = mp.keyValue(i);
                var cmt = (val && val.comment) ? String(val.comment).toLowerCase() : "";
                if (cutIn  === null && cmt === "cut in")  cutIn  = mp.keyTime(i);
                if (cutOut === null && cmt === "cut out") cutOut = mp.keyTime(i);
            }
        }
    } catch (eRead) {}

    if (cutIn === null || cutOut === null) {
        alert("Set Work Area to Cut Markers: couldn't find both \"cut in\" and \"cut out\" composition markers on " + comp.name + ".");
        return;
    }
    if (cutOut <= cutIn) {
        alert("Set Work Area to Cut Markers: \"cut out\" marker is at or before \"cut in\" on " + comp.name + ".");
        return;
    }

    app.beginUndoGroup("Set Work Area to Cut Markers");
    try {
        comp.workAreaStart    = cutIn;
        comp.workAreaDuration = cutOut - cutIn;
    } finally {
        app.endUndoGroup();
    }

})();
