/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Invert Transform Effect
 *
 * For each selected layer with a Distort > Transform effect
 * (matchName "ADBE Geometry2"), adds a second Transform effect named
 * "Transform (Inverse)" whose Anchor / Position / Rotation / Scale /
 * Skew properties are expression-linked to the source and mathematically
 * invert it. Applied in series with the original, the composite is the
 * identity — useful for un-stabilize, composite pre/post comparisons,
 * or parking an element in original-space while the layer above stays
 * transformed.
 *
 * Math (derivation): the Transform effect maps X → Y = S·R·(X − A) + P
 * with anchor A, position P, rotation R, scale S. The inverse is a
 * Transform effect with:
 *   A' = P   (original position becomes new anchor)
 *   P' = A   (original anchor becomes new position)
 *   R' = −R
 *   S' = 1/S (AE percentages: 10000 / s)
 *
 * Skew is inverted as −skew on the same axis. Opacity is left at 100%
 * (can't fully recover < 100% via composition). Uniform Scale checkbox
 * is linked so the inverse matches the source's uniform-vs-independent
 * mode automatically.
 *
 * The source effect is renamed to "Transform (Source)" on first run so
 * the expression references stay stable. Re-running on a layer that
 * already has a "Transform (Inverse)" skips it.
 *
 * Single undo step.
 */

(function () {

    var comp = app.project.activeItem;
    if (!(comp instanceof CompItem)) {
        alert("Invert Transform: please open a composition first.");
        return;
    }

    var sel = comp.selectedLayers;
    if (sel.length === 0) {
        alert("Invert Transform: select at least one layer with a Transform effect.");
        return;
    }

    var XFORM_MATCH = "ADBE Geometry2";
    var SRC_NAME    = "Transform (Source)";
    var INV_NAME    = "Transform (Inverse)";

    function findSourceTransform(layer) {
        try {
            var fx = layer.property("ADBE Effect Parade");
            if (!fx || fx.numProperties === 0) return null;
            for (var i = 1; i <= fx.numProperties; i++) {
                var e = fx.property(i);
                if (e.matchName === XFORM_MATCH && e.name !== INV_NAME) return e;
            }
        } catch (eSF) {}
        return null;
    }

    function hasInverse(layer) {
        try {
            var fx = layer.property("ADBE Effect Parade");
            for (var i = 1; i <= fx.numProperties; i++) {
                if (fx.property(i).name === INV_NAME) return true;
            }
        } catch (e) {}
        return false;
    }

    function setExpr(fx, propName, expr) {
        try {
            var p = fx.property(propName);
            if (p) p.expression = expr;
        } catch (e) {}
    }

    var processed = 0;
    var skippedNoFx = 0;
    var skippedAlready = 0;

    app.beginUndoGroup("Invert Transform Effect");

    try {
        for (var s = 0; s < sel.length; s++) {
            var layer = sel[s];
            var src = findSourceTransform(layer);
            if (!src) { skippedNoFx++; continue; }
            if (hasInverse(layer)) { skippedAlready++; continue; }

            // Stable reference name on the source effect.
            try { src.name = SRC_NAME; } catch (eRn) {}

            // Add the inverse effect (appended to the end of the Effects stack).
            var inv = layer.property("ADBE Effect Parade").addProperty(XFORM_MATCH);
            try { inv.name = INV_NAME; } catch (eRn2) {}

            // Expression helpers. Use explicit .property("name") lookup so
            // AE's expression parser treats the property name as an opaque
            // string — the shorthand effect("X")("Y") form can misroute on
            // grouped properties like "Scale" (which is a Group containing
            // Scale Width and Scale Height in ADBE Geometry2).
            var ref = "effect(\"" + SRC_NAME + "\")";
            var P   = function (name) { return ref + ".property(\"" + name + "\")"; };

            // Swap anchor <-> position (so the inverse pivots around the
            // source's post-transform location and re-centres on the source's
            // anchor).
            setExpr(inv, "Anchor Point",  P("Position"));
            setExpr(inv, "Position",      P("Anchor Point"));

            // Rotation + skew negate on the same axis.
            setExpr(inv, "Rotation",      "-" + P("Rotation"));
            setExpr(inv, "Skew",          "-" + P("Skew"));
            setExpr(inv, "Skew Axis",     P("Skew Axis"));

            // Scale reciprocal — set Scale Width and Scale Height
            // individually. Do NOT set on "Scale" (that's a Property Group
            // in ADBE Geometry2 and setting expression on the group cascades
            // to both children with a broken reference to the source's
            // "Scale" group). Uniform Scale links through so the mode
            // follows the source.
            setExpr(inv, "Uniform Scale", P("Uniform Scale"));
            setExpr(inv, "Scale Width",   "10000 / " + P("Scale Width"));
            setExpr(inv, "Scale Height",  "10000 / " + P("Scale Height"));

            processed++;
        }
    } finally {
        app.endUndoGroup();
    }

    var parts = [];
    parts.push(processed + " layer" + (processed === 1 ? "" : "s") + " inverted");
    if (skippedNoFx > 0)    parts.push(skippedNoFx    + " skipped (no Transform effect)");
    if (skippedAlready > 0) parts.push(skippedAlready + " skipped (already has an Inverse)");
    alert("Invert Transform: " + parts.join(", ") + ".");

})();
