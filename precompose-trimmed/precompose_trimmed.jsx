/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Precompose Trimmed
 *
 * Each selected layer is precomposed individually with "move all attributes".
 * The new precomp gets the full parent comp duration, then the precomp layer
 * is trimmed back to the original layer's in/out points.
 *
 * Multiple layers → each gets its own precomp, processed highest-index first
 * so earlier index-shifts never corrupt the remaining layers.
 *
 * Single undo step.
 */

(function () {

    // ── validate ───────────────────────────────────────────────────────────────

    var comp = app.project.activeItem;
    if (!(comp instanceof CompItem)) {
        alert("Precompose Trimmed: please open a composition first.");
        return;
    }

    var selected = comp.selectedLayers;
    if (selected.length === 0) {
        alert("Precompose Trimmed: select at least one layer.");
        return;
    }

    // ── snapshot all layer data before touching anything ───────────────────────

    var layerInfos = [];
    for (var i = 0; i < selected.length; i++) {
        var layer = selected[i];
        layerInfos.push({
            index:    layer.index,
            name:     layer.name,
            inPoint:  layer.inPoint,
            outPoint: layer.outPoint
        });
    }

    // Process highest index first — precomposing a lower layer in the stack
    // doesn't shift the indices of layers above it, so earlier passes stay valid.
    layerInfos.sort(function (a, b) { return b.index - a.index; });

    // ── helpers ────────────────────────────────────────────────────────────────

    function nameExists(name) {
        for (var k = 1; k <= app.project.numItems; k++) {
            if (app.project.item(k).name === name) return true;
        }
        return false;
    }

    function uniqueName(base) {
        var name = base;
        var suffix = 1;
        while (nameExists(name)) {
            name = base + "_" + suffix;
            suffix++;
        }
        return name;
    }

    // ── loop ───────────────────────────────────────────────────────────────────

    app.beginUndoGroup("Precompose Trimmed");

    try {
        for (var i = 0; i < layerInfos.length; i++) {
            var info = layerInfos[i];

            var precompName = uniqueName(info.name + "_precomp");

            // true = "Move all attributes into the new composition"
            var newComp = comp.layers.precompose([info.index], precompName, true);

            // Full parent comp length — discard AE's automatic duration trim.
            newComp.duration = comp.duration;

            // Find the replacement layer in the parent comp.
            var precompLayer = null;
            for (var j = 1; j <= comp.numLayers; j++) {
                if (comp.layer(j).source === newComp) {
                    precompLayer = comp.layer(j);
                    break;
                }
            }

            if (!precompLayer) {
                alert("Precompose Trimmed: could not locate new precomp layer for \"" + info.name + "\".");
                return;
            }

            precompLayer.inPoint  = info.inPoint;
            precompLayer.outPoint = info.outPoint;
        }

    } finally {
        app.endUndoGroup();
    }

})();
