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

    // Snapshot non-target selection by index. precompose() collapses
    // comp.selectedLayers to just the most-recently-created precomp layer,
    // so anything the user had selected that wasn't a target would silently
    // drop out. After the loop we re-select those originals plus every new
    // precomp layer we just created. Matches the roundtrip's autoPrecomposeTrimmed.
    var targetIndices = {};
    for (var ti = 0; ti < layerInfos.length; ti++) targetIndices[layerInfos[ti].index] = true;
    var preservedSelection = [];
    for (var pi = 1; pi <= comp.numLayers; pi++) {
        var pl = comp.layer(pi);
        if (pl.selected && !targetIndices[pi]) preservedSelection.push(pl);
    }
    var newPrecompLayers = [];

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
            newPrecompLayers.push(precompLayer);
        }

        // Restore the combined selection: preserved originals + every new precomp.
        try {
            for (var dsel = 1; dsel <= comp.numLayers; dsel++) comp.layer(dsel).selected = false;
            for (var psi = 0; psi < preservedSelection.length; psi++) {
                try { preservedSelection[psi].selected = true; } catch(ePs) {}
            }
            for (var npi = 0; npi < newPrecompLayers.length; npi++) {
                try { newPrecompLayers[npi].selected = true; } catch(eNp) {}
            }
        } catch(eSelPC) {}

    } finally {
        app.endUndoGroup();
    }

})();
