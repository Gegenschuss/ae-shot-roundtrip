/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

// ─────────────────────────────────────────────────────────────────────────────
// batch_syntheyes_to_ae.jsx
//
// Batch-runs SynthEyes-exported .jsx files in After Effects and saves each
// result as a .aep project next to the original .jsx file.
//
// HOW TO USE:
//   In After Effects: File > Scripts > Run Script File… > pick this file
//   Or place it in your AE Scripts folder to access from the Scripts menu.
// ─────────────────────────────────────────────────────────────────────────────

(function batchSyntheyesToAE() {

    // ── Settings ──────────────────────────────────────────────────────────────
    var SKIP_EXISTING = true;    // true  = skip .jsx files that already have a .aep
                                 // false = overwrite existing .aep files
    var SEARCH_SUBFOLDERS = false; // true = also process .jsx files in subfolders

    // ── Pick folder ───────────────────────────────────────────────────────────
    var folder = Folder.selectDialog("Select folder containing SynthEyes .jsx files");
    if (!folder) return; // user cancelled

    // ── Collect .jsx files ────────────────────────────────────────────────────
    var thisScriptPath = (new File($.fileName)).fsName; // exclude self

    function collectJSX(dir, recursive) {
        var results = [];
        var items = dir.getFiles();
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item instanceof File && /\.jsx$/i.test(item.name)) {
                if (item.fsName !== thisScriptPath) {
                    results.push(item);
                }
            } else if (recursive && item instanceof Folder) {
                results = results.concat(collectJSX(item, true));
            }
        }
        // Sort alphabetically for predictable order
        results.sort(function(a, b) {
            return a.fsName < b.fsName ? -1 : a.fsName > b.fsName ? 1 : 0;
        });
        return results;
    }

    var jsxFiles = collectJSX(folder, SEARCH_SUBFOLDERS);

    if (jsxFiles.length === 0) {
        alert("No .jsx files found in:\n" + folder.fsName);
        return;
    }

    // ── Confirm ───────────────────────────────────────────────────────────────
    var skipCount = 0;
    if (SKIP_EXISTING) {
        for (var k = 0; k < jsxFiles.length; k++) {
            var aepCheck = new File(jsxFiles[k].fsName.replace(/\.jsx$/i, ".aep"));
            if (aepCheck.exists) skipCount++;
        }
    }

    var confirmMsg = "Found " + jsxFiles.length + " .jsx file(s) in:\n" + folder.fsName;
    if (skipCount > 0) confirmMsg += "\n\n(" + skipCount + " already have a .aep and will be skipped)";
    confirmMsg += "\n\nProceed?";
    if (!confirm(confirmMsg)) return;

    // ── Process each file ─────────────────────────────────────────────────────
    var processed = 0;
    var skipped   = 0;
    var errors    = 0;
    var errorLog  = [];

    for (var i = 0; i < jsxFiles.length; i++) {
        var jsxFile = jsxFiles[i];
        var aepPath = jsxFile.fsName.replace(/\.jsx$/i, ".aep");
        var aepFile = new File(aepPath);

        // Skip if output already exists
        if (SKIP_EXISTING && aepFile.exists) {
            skipped++;
            continue;
        }

        try {
            // Run the SynthEyes script — it typically calls app.newProject() internally
            $.evalFile(jsxFile);

            // Save the resulting project as .aep next to the .jsx
            app.project.save(aepFile);

            // Close the project cleanly (already saved, no prompt needed)
            app.project.close(CloseOptions.DO_NOT_SAVE_CHANGES);

            processed++;

        } catch (e) {
            errors++;
            errorLog.push("• " + jsxFile.name + "\n  " + e.message);
        }
    }

    // ── Summary ───────────────────────────────────────────────────────────────
    var summary = "Batch complete!\n\n"
        + "✓ Processed : " + processed + "\n"
        + "– Skipped   : " + skipped   + " (already had .aep)\n"
        + "✗ Errors    : " + errors;

    if (errorLog.length > 0) {
        summary += "\n\nError details:\n" + errorLog.join("\n\n");
    }

    alert(summary);

})();
