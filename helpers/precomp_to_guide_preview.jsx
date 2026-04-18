/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Precomp to Guide Preview
 *
 * For each selected layer:
 *   1. Precomposes with "leave all attributes" (transforms stay on the parent layer).
 *   2. Scales the parent-comp layer by 2× so the final size is unchanged.
 *   3. Inside the new precomp, marks the original footage layer as a guide
 *      layer and halves its scale.
 *   4. Halves the new precomp's width and height (e.g. 4K UHD → HD).
 *
 * Net effect: a half-resolution working copy for faster previews, displayed
 * at the original size in the parent comp, with the footage flagged as a
 * guide layer so it doesn't render.
 *
 * Multi-layer selections → each layer gets its own _preview precomp,
 * processed highest-index first so earlier passes don't invalidate the rest.
 *
 * Single undo step.
 */

(function () {

    var comp = app.project.activeItem;
    if (!(comp instanceof CompItem)) {
        alert("Precomp to Guide Preview: please open a composition first.");
        return;
    }

    var selected = comp.selectedLayers;
    if (selected.length === 0) {
        alert("Precomp to Guide Preview: select at least one layer.");
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

    app.beginUndoGroup("Precomp to Guide Preview");

    try {
        for (var li = 0; li < layerInfos.length; li++) {
            var info      = layerInfos[li];
            // Strip AE's auto-generated ".mov Comp N" suffix from the source
            // name so the new precomp reads "<shot>_downrez" instead of
            // "<shot>.mov Comp 1_downrez".
            var baseName  = info.name.replace(/\.mov Comp \d+/i, "");
            var pcName    = uniqueName(baseName + "_downrez");

            // false = "Leave all attributes in the source layer"
            var newComp = comp.layers.precompose([info.index], pcName, false);

            // Locate the precomp layer in the parent comp.
            var pcLayer = null;
            for (var j = 1; j <= comp.numLayers; j++) {
                if (comp.layer(j).source === newComp) { pcLayer = comp.layer(j); break; }
            }
            if (!pcLayer) {
                alert("Precomp to Guide Preview: could not locate new precomp layer for \"" + info.name + "\".");
                continue;
            }

            // 1. Parent-comp layer: multiply scale by 2 (compensates for the upcoming half-res precomp).
            try {
                var parentScale = pcLayer.property("ADBE Transform Group").property("ADBE Scale");
                var ps = parentScale.value;
                parentScale.setValue([ps[0] * 2, ps[1] * 2]);
            } catch (eParentScale) {}

            // 2. Inside: mark the (single) inner layer as guide + halve scale + halve position.
            if (newComp.numLayers >= 1) {
                var inner = newComp.layer(1);
                try { inner.guideLayer = true; } catch (eGuide) {}
                try {
                    var innerTransform = inner.property("ADBE Transform Group");

                    var innerScale = innerTransform.property("ADBE Scale");
                    var is = innerScale.value;
                    innerScale.setValue([is[0] * 0.5, is[1] * 0.5]);

                    var innerPos = innerTransform.property("ADBE Position");
                    var ip = innerPos.value;
                    if (ip.length === 3) {
                        innerPos.setValue([ip[0] * 0.5, ip[1] * 0.5, ip[2]]); // leave Z alone
                    } else {
                        innerPos.setValue([ip[0] * 0.5, ip[1] * 0.5]);
                    }
                } catch (eInnerXform) {}
            }

            // 3. Halve the precomp's dimensions. Force even so codecs stay happy.
            var newW = Math.max(2, Math.floor(newComp.width  / 2));
            var newH = Math.max(2, Math.floor(newComp.height / 2));
            if (newW % 2 !== 0) newW--;
            if (newH % 2 !== 0) newH--;
            newComp.width  = newW;
            newComp.height = newH;
        }
    } finally {
        app.endUndoGroup();
    }

})();
