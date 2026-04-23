/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Oversize Precomp 2×
 *
 * For each selected layer:
 *   1. Precomposes with "leave all attributes" (transforms stay on the parent layer).
 *   2. Parent-comp precomp layer stays at 100% scale (unchanged by this script).
 *   3. Inside the new precomp, doubles the inner layer's scale and position so
 *      it stays centred after the upcoming 2× canvas expansion.
 *   4. Doubles the new precomp's width and height.
 *
 * Net effect: the precomp is twice as large in every dimension and displays
 * at 2× in the parent comp. Useful for zoom-in / overscan workflows where
 * you want the higher-resolution canvas available without an outer scale
 * compensation.
 *
 * Multi-layer selections → each layer gets its own _oversize precomp,
 * processed highest-index first so earlier passes don't invalidate the rest.
 *
 * Single undo step.
 */

(function () {

    var comp = app.project.activeItem;
    if (!(comp instanceof CompItem)) {
        alert("Oversize Precomp 2×: please open a composition first.");
        return;
    }

    var selected = comp.selectedLayers;
    if (selected.length === 0) {
        alert("Oversize Precomp 2×: select at least one layer.");
        return;
    }

    // Snapshot before touching anything.
    var layerInfos = [];
    for (var i = 0; i < selected.length; i++) {
        layerInfos.push({ index: selected[i].index, name: selected[i].name });
    }

    // Highest index first — precomposing a lower layer doesn't shift higher indices.
    layerInfos.sort(function (a, b) { return b.index - a.index; });

    function nameExists(name) {
        for (var k = 1; k <= app.project.numItems; k++) {
            if (app.project.item(k).name === name) return true;
        }
        return false;
    }

    function uniqueName(base) {
        var n = base, suffix = 1;
        while (nameExists(n)) { n = base + "_" + (suffix++); }
        return n;
    }

    app.beginUndoGroup("Oversize Precomp 2×");

    try {
        for (var li = 0; li < layerInfos.length; li++) {
            var info      = layerInfos[li];
            // Strip AE's auto-generated ".mov Comp N" suffix from the source
            // name so the new precomp reads "<shot>_oversize" instead of
            // "<shot>.mov Comp 1_oversize".
            var baseName  = info.name.replace(/\.mov Comp \d+/i, "");
            var pcName    = uniqueName(baseName + "_oversize");

            // false = "Leave all attributes in the source layer"
            var newComp = comp.layers.precompose([info.index], pcName, false);

            // Locate the precomp layer in the parent comp.
            var pcLayer = null;
            for (var j = 1; j <= comp.numLayers; j++) {
                if (comp.layer(j).source === newComp) { pcLayer = comp.layer(j); break; }
            }
            if (!pcLayer) {
                alert("Oversize Precomp 2×: could not locate new precomp layer for \"" + info.name + "\".");
                continue;
            }

            // 1. Parent-comp precomp layer stays at 100% scale — leave it
            //    alone. The 2× precomp will show 2× larger in the parent
            //    comp as a deliberate result.

            // 2. Inside: double the (single) inner layer's scale + position so
            //    it stays centred once the precomp canvas expands.
            if (newComp.numLayers >= 1) {
                var inner = newComp.layer(1);
                try {
                    var innerTransform = inner.property("ADBE Transform Group");

                    var innerScale = innerTransform.property("ADBE Scale");
                    var is = innerScale.value;
                    innerScale.setValue([is[0] * 2, is[1] * 2]);

                    var innerPos = innerTransform.property("ADBE Position");
                    var ip = innerPos.value;
                    if (ip.length === 3) {
                        innerPos.setValue([ip[0] * 2, ip[1] * 2, ip[2]]); // leave Z alone
                    } else {
                        innerPos.setValue([ip[0] * 2, ip[1] * 2]);
                    }
                } catch (eInnerXform) {}
            }

            // 3. Double the precomp's dimensions. Force even so codecs stay happy.
            var newW = newComp.width  * 2;
            var newH = newComp.height * 2;
            if (newW % 2 !== 0) newW--;
            if (newH % 2 !== 0) newH--;
            newComp.width  = newW;
            newComp.height = newH;
        }
    } finally {
        app.endUndoGroup();
    }

})();
