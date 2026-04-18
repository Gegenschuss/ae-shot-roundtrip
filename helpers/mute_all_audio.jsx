/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Mute All Audio (recursive)
 *
 * For every selected layer in the active comp, disables its audio switch
 * AND walks into any precomp source, muting audio on every nested layer
 * all the way down.
 *
 * Shared precomps are visited only once so walk time stays bounded.
 *
 * Single undo step.
 */

(function () {

    var comp = app.project.activeItem;
    if (!(comp instanceof CompItem)) {
        alert("Mute All Audio: please open a composition first.");
        return;
    }

    var selected = comp.selectedLayers;
    if (selected.length === 0) {
        alert("Mute All Audio: select at least one layer.");
        return;
    }

    var muted   = 0;
    var visited = {}; // comp.id → true, avoids revisiting shared precomps

    function muteLayer(layer) {
        try {
            if (layer.hasAudio && layer.audioEnabled) {
                layer.audioEnabled = false;
                muted++;
            }
        } catch (e) {
            // Layer types without an audio switch (nulls, cameras, lights) throw — skip.
        }
    }

    function walkComp(c) {
        if (!(c instanceof CompItem)) return;
        if (visited[c.id]) return;
        visited[c.id] = true;
        for (var i = 1; i <= c.numLayers; i++) {
            var inner = c.layer(i);
            muteLayer(inner);
            if (inner.source instanceof CompItem) walkComp(inner.source);
        }
    }

    app.beginUndoGroup("Mute All Audio (recursive)");

    try {
        for (var s = 0; s < selected.length; s++) {
            var layer = selected[s];
            muteLayer(layer);
            if (layer.source instanceof CompItem) walkComp(layer.source);
        }
    } finally {
        app.endUndoGroup();
    }

    alert("Mute All Audio: " + muted + " layer(s) muted.");

})();
