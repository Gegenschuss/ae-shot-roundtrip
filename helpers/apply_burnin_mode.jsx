/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/*
================================================================================
  APPLY BURNIN MODE
  After Effects ExtendScript
================================================================================

Reads the "Burnin CTRL" null in the active comp (mainComp) and pushes the
"Render" checkbox down to every shot _comp's Guide Burnin layer by
flipping its guideLayer flag:

  Render = 0  →  guideLayer = true   (visible in AE viewer, excluded from render)
  Render = 1  →  guideLayer = false  (visible AND rendered)

The "Show" checkbox on the same CTRL controls opacity via a cross-comp
expression baked in by Shot Roundtrip; it's live and doesn't need this
helper to run.

Shot Roundtrip creates the CTRL null automatically. If it's missing,
this script exits with a prompt to re-run Shot Roundtrip.
================================================================================
*/

(function () {
    var proj = app.project;
    if (!proj) { alert("Apply Burnin Mode: no project open."); return; }

    var mainComp = proj.activeItem;
    if (!(mainComp instanceof CompItem)) {
        alert("Apply Burnin Mode: open the main comp that holds the Burnin CTRL null.");
        return;
    }

    var ctrl = mainComp.layers.byName("Burnin CTRL");
    if (!ctrl) {
        alert("Apply Burnin Mode: no \"Burnin CTRL\" null in " + mainComp.name
            + ".\nRe-run Shot Roundtrip to create it.");
        return;
    }

    var renderVal = 0;
    try {
        var rFx = ctrl.Effects.property("Render");
        if (rFx) renderVal = rFx.property(1).value;
    } catch (eRV) {
        alert("Apply Burnin Mode: could not read \"Render\" checkbox on Burnin CTRL.\n" + eRV.toString());
        return;
    }
    var renderOn = (renderVal === 1 || renderVal === true);

    // Iterate every *_comp / *_comp_OS in the project and flip Guide Burnin.
    var touched = 0, missing = 0;
    app.beginUndoGroup("Apply Burnin Mode");
    try {
        for (var i = 1; i <= proj.numItems; i++) {
            var item = proj.item(i);
            if (!(item instanceof CompItem)) continue;
            if (!/_comp(_OS)?$/i.test(item.name)) continue;

            var gl = item.layers.byName("Guide Burnin");
            if (!gl) { missing++; continue; }
            try {
                gl.locked     = false;
                gl.guideLayer = !renderOn;
                gl.enabled    = true;
                gl.locked     = true;
                touched++;
            } catch (eGL) { /* layer may be locked or in a weird state; skip */ }
        }
    } finally {
        app.endUndoGroup();
    }

    alert("Apply Burnin Mode\n\n"
        + "Render flag: " + (renderOn ? "ON (Guide Burnin will render)" : "OFF (guide only, not rendered)") + "\n"
        + "Shot comps updated: " + touched
        + (missing > 0 ? ("\nShot comps with no Guide Burnin: " + missing) : ""));
})();
