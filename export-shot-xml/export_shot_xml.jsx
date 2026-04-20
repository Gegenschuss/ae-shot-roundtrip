/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/**
 * Export Shot XML for DaVinci Resolve  –  FCP7 / xmeml v4
 *
 * Scans the "Shots" folder for every "*_comp" composition, picks the
 * "active" footage layer for each shot, and writes an FCP7-compatible
 * XML matching Premiere Pro's export format, which DaVinci Resolve
 * imports reliably.
 *
 * Two layouts are supported (matching Import Returns):
 *
 *   1. FLAT — footage sits directly in {shot}_comp. We pick the topmost
 *      enabled footage layer whose name starts with the shot number
 *      (e.g. "KM_050_plate_v02.mov" in comp "KM_050_comp").
 *
 *   2. PRECOMP — {shot}_comp contains a {shot}_stack precomp created by
 *      Shot Roundtrip. Inside that precomp, the stack order is
 *      grade (top) → render → plate (bottom), and the importer disables
 *      older files within each category. We recurse into the precomp
 *      and pick the topmost enabled footage layer there — naturally the
 *      most-finished visible version (latest grade if present, else the
 *      latest VFX render, else the plate).
 *
 * Import in DaVinci Resolve via:
 *   File → Import Timeline → Import AAF, EDL, XML…
 */

(function () {

    // ─── helpers ────────────────────────────────────────────────────────────

    function isFootageLayer(layer) {
        if (!(layer instanceof AVLayer)) return false;
        if (!layer.source) return false;
        if (!(layer.source instanceof FootageItem)) return false;
        if (!layer.source.file) return false;
        return true;
    }

    function xmlEscape(str) {
        return String(str)
            .replace(/&/g,  "&amp;")
            .replace(/</g,  "&lt;")
            .replace(/>/g,  "&gt;")
            .replace(/"/g,  "&quot;")
            .replace(/'/g,  "&apos;");
    }

    /**
     * Convert an OS path to a file://localhost/ URL.
     * Premiere / Resolve requires this exact form (not just file://).
     */
    function toFileURL(fsName) {
        var p = fsName.replace(/\\/g, "/");
        if (p.charAt(0) !== "/") p = "/" + p;   // Windows: /C:/path
        p = p.replace(/ /g,  "%20")
              .replace(/\[/g, "%5B")
              .replace(/\]/g, "%5D")
              .replace(/\(/g, "%28")
              .replace(/\)/g, "%29")
              .replace(/\+/g, "%2B")
              .replace(/,/g,  "%2C");
        return "file://localhost" + p;
    }

    function pad2(n) { return n < 10 ? "0" + n : String(n); }

    /** Format frames as HH:MM:SS:FF timecode. */
    function framesToTC(totalFrames, fps) {
        var r = Math.round(fps);
        var ff = totalFrames % r;
        var ss = Math.floor(totalFrames / r) % 60;
        var mm = Math.floor(totalFrames / (r * 60)) % 60;
        var hh = Math.floor(totalFrames / (r * 3600));
        return pad2(hh) + ":" + pad2(mm) + ":" + pad2(ss) + ":" + pad2(ff);
    }

    /**
     * Premiere ticks: 254016000000 ticks per second.
     * Returns a string to avoid JS integer precision issues.
     */
    function pproTicks(frames, fps) {
        // frames / fps * 254016000000 — keep integer by doing frames * (254016000000 / fps)
        var ticksPerFrame = 254016000000 / Math.round(fps);
        return String(Math.round(frames * ticksPerFrame));
    }

    /**
     * Read a 4-byte big-endian unsigned integer from a binary string.
     * Uses multiplication instead of << to avoid JS signed-int overflow.
     */
    function readU32(str, off) {
        return ((str.charCodeAt(off)     & 0xFF) * 16777216)
             + ((str.charCodeAt(off + 1) & 0xFF) * 65536)
             + ((str.charCodeAt(off + 2) & 0xFF) * 256)
             +  (str.charCodeAt(off + 3) & 0xFF);
    }

    /**
     * Parse the QuickTime mvhd atom from a raw binary string.
     * Returns { fps, totalFrames } if the timescale is a standard video fps,
     * otherwise returns null.
     *
     * mvhd v0: [type(4)] version(1) flags(3) creation(4) modification(4) timescale(4) duration(4)
     * mvhd v1: [type(4)] version(1) flags(3) creation(8) modification(8) timescale(4) duration(8)
     */
    function parseMvhd(raw) {
        var idx = raw.indexOf("mvhd");
        if (idx < 0) return null;
        var off     = idx + 4;
        var version = raw.charCodeAt(off) & 0xFF;
        var tsOff   = off + (version === 0 ? 12 : 20);
        var durOff  = tsOff + 4;

        var ts  = readU32(raw, tsOff);
        var dur = (version === 0)
                    ? readU32(raw, durOff)
                    : readU32(raw, durOff + 4); // lower 32 bits of 64-bit duration

        var validFps = { 24:1, 25:1, 30:1, 48:1, 50:1, 60:1 };
        if (!validFps[ts]) return null;

        return { fps: ts, totalFrames: dur };  // dur == frames when timescale == fps
    }

    /**
     * Read the embedded start timecode (in frames) from a QuickTime/MOV file.
     *
     * The TC is stored as a 4-byte big-endian frame count in the tmcd track's
     * sample data.  Strategy:
     *   1. Find the 'hdlr' atom whose handler_type == 'tmcd'
     *   2. Find the next 'stco' atom (chunk-offset table for that track)
     *   3. Read the first chunk offset → seek to it → read 4 bytes = TC frame
     *
     * Returns 0 if the TC cannot be determined.
     */
    function readFileTCFrame(file, raw) {
        try {
            // Find hdlr with handler_type 'tmcd'
            // hdlr atom layout (from the 'hdlr' type tag):
            //   +0  "hdlr"         type (4)
            //   +4  version+flags  (4)
            //   +8  pre_defined    (4)
            //   +12 handler_type   (4) ← "tmcd", "vide", "soun" …
            var searchFrom = 0;
            var hdlrIdx = -1;
            while (true) {
                var hi = raw.indexOf("hdlr", searchFrom);
                if (hi < 0) break;
                if (raw.substr(hi + 12, 4) === "tmcd") { hdlrIdx = hi; break; }
                searchFrom = hi + 1;
            }
            if (hdlrIdx < 0) return 0;

            // Find the chunk-offset atom that belongs to the tmcd track.
            // The tmcd track has exactly 1 sample, so its stco/co64 has entry_count == 1.
            // We scan forward from the tmcd hdlr and accept the first stco or co64 whose
            // entry_count equals 1.  This skips false-positive matches (random bytes that
            // spell "stco"/"co64" in mdat data, or other tracks' offset tables which have
            // many entries).
            var chunkOff   = -1;
            var chunkOffHi = 0;   // high 32 bits of co64 offset (0 for stco / small files)

            // ── stco (32-bit offsets) ──────────────────────────────────────
            var sfrom = hdlrIdx;
            while (chunkOff < 0) {
                var si = raw.indexOf("stco", sfrom);
                if (si < 0) break;
                if (readU32(raw, si + 8) === 1) {          // entry_count == 1
                    chunkOff = readU32(raw, si + 12);
                    break;
                }
                sfrom = si + 1;
            }

            // ── co64 (64-bit offsets, fallback) ───────────────────────────
            if (chunkOff < 0) {
                sfrom = hdlrIdx;
                while (chunkOff < 0) {
                    var ci = raw.indexOf("co64", sfrom);
                    if (ci < 0) break;
                    if (readU32(raw, ci + 8) === 1) {      // entry_count == 1
                        chunkOffHi = readU32(raw, ci + 12); // high 32 bits
                        chunkOff   = readU32(raw, ci + 16); // low  32 bits
                        break;
                    }
                    sfrom = ci + 1;
                }
            }

            if (chunkOff < 0) return 0;

            // Read the 4-byte TC frame value at chunkOff.
            // chunkOffHi holds the high 32 bits of a co64 entry (0 for stco).
            var tcFrame;
            if (chunkOffHi === 0 && chunkOff + 4 <= raw.length) {
                // Within the already-loaded 4 MB buffer
                tcFrame = readU32(raw, chunkOff);
            } else if (chunkOffHi === 0) {
                // 32-bit offset beyond the buffer — seek in the file
                file.encoding = "BINARY";
                if (!file.open("r")) return 0;
                file.seek(chunkOff);
                var tcRaw = file.read(4);
                file.close();
                tcFrame = readU32(tcRaw, 0);
            } else {
                // 64-bit offset (file > 4 GB) — ExtendScript's File.seek() tops out at
                // ~4 GB, so use a shell dd command to reach the correct byte position.
                var fullOff = chunkOffHi * 4294967296 + chunkOff;
                var safeFs  = file.fsName.replace(/'/g, "'\\''");
                var hex = system.callSystem(
                    "dd if='" + safeFs + "' bs=1 skip=" + fullOff
                    + " count=4 2>/dev/null | od -An -tx1 | tr -d ' \\n'"
                );
                if (!hex || hex.length < 8) return 0;
                hex = hex.replace(/[^0-9a-fA-F]/g, "");
                if (hex.length < 8) return 0;
                tcFrame = parseInt(hex.substr(0, 8), 16);
            }
            // Sanity: reject values > 24 h at 60 fps (= 5 184 000 frames).
            return (tcFrame > 0 && tcFrame < 5184000) ? tcFrame : 0;
        } catch (e) {
            return 0;
        }
    }

    /**
     * Read native fps, total frame count, and embedded start TC frame from a
     * QuickTime/MOV file.  Returns { fps, totalFrames, tcStartFrame } or null.
     *
     * Reading 4 MB covers the moov atom for virtually all ProRes files.
     */
    function getFileInfo(file) {
        try {
            file.encoding = "BINARY";
            if (!file.open("r")) return null;
            var raw = file.read(4194304);
            file.close();

            var mvhd = parseMvhd(raw);
            if (!mvhd) return null;

            var tcFrame = readFileTCFrame(file, raw);
            return { fps: mvhd.fps, totalFrames: mvhd.totalFrames, tcStartFrame: tcFrame };
        } catch (e) {
            return null;
        }
    }

    /** FCP7 <rate> block. */
    function rateBlock(fps, indent) {
        var sp = indent || "";
        var ntsc = (Math.abs(fps - 23.976) < 0.01 ||
                    Math.abs(fps - 29.97)  < 0.01 ||
                    Math.abs(fps - 59.94)  < 0.01) ? "TRUE" : "FALSE";
        return sp + "<rate>\n"
             + sp + "\t<timebase>" + Math.round(fps) + "</timebase>\n"
             + sp + "\t<ntsc>"     + ntsc            + "</ntsc>\n"
             + sp + "</rate>";
    }

    function collectCompItems(folder, result) {
        for (var i = 1; i <= folder.numItems; i++) {
            var item = folder.item(i);
            if (item instanceof FolderItem) {
                collectCompItems(item, result);
            } else if (item instanceof CompItem) {
                if (/[_\s]comp(_OS)?$/i.test(item.name)) result.push(item);
            }
        }
    }

    // Finds the "Shots" folder at project ROOT only. Projects sometimes
    // contain additional folders named "Shots" nested elsewhere (per-shot
    // subfolders, leftovers from other workflows, etc.) — those are ignored
    // here so the XML export only looks at the authoritative root-level bin
    // that the roundtrip creates via proj.items.addFolder().
    function findShotsFolder(root) {
        for (var i = 1; i <= root.numItems; i++) {
            var item = root.item(i);
            if (item instanceof FolderItem && item.name.toLowerCase() === "shots") {
                return item;
            }
        }
        return null;
    }

    function shotNameFromComp(compName) {
        // Must mirror the leniency of collectCompItems() so "KM_010_comp"
        // and "KM_010_comp_OS" both resolve to "KM_010" — otherwise the
        // shot-prefix regex in pickFootageLayer silently fails to match
        // any layer inside overscan variants.
        return compName.replace(/[_\s]comp(_OS)?$/i, "");
    }

    /**
     * Return the topmost enabled footage layer inside a precomp. No
     * shot-name filter — by construction the {shot}_stack precomp only
     * holds plate/render/grade variants of this shot. Import Returns
     * keeps the stack ordered grade → render → plate top-to-bottom
     * and disables older files within each category, so "topmost enabled
     * footage" resolves to the most-finished visible version.
     */
    function pickTopmostFootageLayerInPrecomp(comp) {
        for (var i = 1; i <= comp.numLayers; i++) {
            var layer = comp.layer(i);
            if (!layer.enabled) continue;
            if (!isFootageLayer(layer)) continue;
            return layer;
        }
        return null;
    }

    /**
     * Return the active footage layer for a shot comp. Supports both
     * layouts written by Import Returns:
     *
     *   - PRECOMP: a {shot}_stack precomp layer lives directly in _comp
     *     as the hero. We recurse into it and pick the topmost enabled
     *     footage layer inside — naturally the newest grade, else the
     *     newest VFX render, else the plate itself.
     *   - FLAT: footage sits directly in _comp. We pick the topmost
     *     enabled footage layer whose name starts with the shot number
     *     (e.g. "KM_050_plate_v02.mov" or "KM_050_render.[####].exr" in
     *     comp "KM_050_comp").
     *
     * Accepts any footage item — movie files (mov, mp4, …) and image
     * sequences (dpx, exr, tif, …) alike.
     */
    function pickFootageLayer(comp) {
        var shot = shotNameFromComp(comp.name);
        var escaped = shot.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        var shotRe = new RegExp("^" + escaped, "i");
        for (var i = 1; i <= comp.numLayers; i++) {
            var layer = comp.layer(i);
            if (!layer.enabled) continue;
            if (!layer.source) continue;
            if (layer.source instanceof CompItem &&
                /_stack(_OS)?$/i.test(layer.source.name) &&
                shotRe.test(layer.source.name)) {
                var inner = pickTopmostFootageLayerInPrecomp(layer.source);
                if (inner) return inner;
                continue;
            }
            if (!isFootageLayer(layer)) continue;
            if (shotRe.test(layer.name)) return layer;
        }
        return null;
    }

    // ─── main ────────────────────────────────────────────────────────────────

    var proj = app.project;
    if (!proj) { alert("No project open."); return; }

    // Project base name used for the sequence name inside the XML and the
    // generated output filename. Strip the .aep extension. Fall back to
    // "Shot Export" only if the project is still untitled.
    var projectBaseName = (proj.file && proj.file.displayName)
                        ? proj.file.displayName.replace(/\.aep$/i, "")
                        : "Shot Export";
    // Sanitise for use in a filename: strip characters that filesystems hate.
    var projectFileStem = projectBaseName.replace(/[\/\\:*?"<>|]/g, "_");

    var shotsFolder = findShotsFolder(proj.rootFolder);
    if (!shotsFolder) {
        alert("Could not find a folder named \"Shots\" in the project.");
        return;
    }

    var comps = [];
    collectCompItems(shotsFolder, comps);
    if (comps.length === 0) {
        alert("No \"*_comp\" compositions found inside the Shots folder.");
        return;
    }

    comps.sort(function (a, b) {
        return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
    });

    // ─── gather clip data ────────────────────────────────────────────────────

    var clips    = [];
    var warnings = [];
    var seqFps   = 0;

    for (var c = 0; c < comps.length; c++) {
        var comp  = comps[c];
        var layer = pickFootageLayer(comp);

        if (!layer) {
            warnings.push(comp.name + ": no active footage layer — skipped.");
            continue;
        }

        var compFps = comp.frameRate;
        if (compFps > seqFps) seqFps = compFps;

        // Read native fps, total frame count, and embedded TC start frame directly
        // from the file's moov/mvhd atom — immune to AE's "Interpret Footage"
        // settings and source.duration inaccuracies on files with non-zero TC.
        // Only QuickTime-family files have a moov atom; image sequences and
        // other formats go straight to the AE fallback.
        var isQuickTime = /\.(mov|mp4|m4v|qt)$/i.test(layer.source.file.fsName);
        var fileInfo = isQuickTime ? getFileInfo(layer.source.file) : null;
        var srcFps = (fileInfo && fileInfo.fps > 0) ? fileInfo.fps
                   : (layer.source.frameRate > 0)   ? layer.source.frameRate
                   : compFps;

        // Total file duration from the mvhd atom (AE's source.duration can be
        // wrong when the file has a non-zero embedded start timecode).
        var totalDurF = (fileInfo && fileInfo.totalFrames > 0)
                          ? fileInfo.totalFrames
                          : Math.round(layer.source.duration * srcFps);

        if (isQuickTime && !fileInfo) {
            warnings.push(comp.name + ": could not read file header for \""
                + layer.source.file.displayName
                + "\" — fps=" + srcFps + " totalDurF=" + totalDurF + " (fallback)");
        }

        // Embedded start TC frame (needed for <file><timecode> so Resolve can
        // locate the clip by its actual TC range, not physical frame 0).
        var fileTCFrame = fileInfo ? fileInfo.tcStartFrame : 0;

        // Full clip: physical frame 0 to end, ignoring the AE layer trim.
        // <file><timecode><frame> = fileTCFrame tells Resolve that physical frame 0
        // corresponds to the embedded TC start, so in=0 correctly maps to TC start.
        var srcInF  = 0;
        var srcOutF = totalDurF;
        var durF    = totalDurF;

        clips.push({
            n:           clips.length + 1,
            name:        layer.source.file.displayName,
            comp:        comp.name,
            path:        layer.source.file.fsName,
            compFps:     compFps,     // used only for seq start/end conversion
            srcFps:      srcFps,      // native fps of the source file
            width:       comp.width,
            height:      comp.height,
            srcInF:      srcInF,      // TC-based in frame
            srcOutF:     srcOutF,     // TC-based out frame
            durF:        durF,
            totalDurF:   totalDurF,   // from file header
            fileTCFrame: fileTCFrame  // embedded TC start frame
        });
    }

    if (clips.length === 0) {
        alert("No clips could be exported.\n\n" + warnings.join("\n"));
        return;
    }

    // ─── sequence-relative start/end (all in seqFps) ─────────────────────────

    var seqHead = 0;
    for (var i = 0; i < clips.length; i++) {
        var cl = clips[i];
        // <start>/<end>: position in the SEQUENCE timeline (seqFps frames).
        // Convert using compFps (how AE timed the clip in the comp).
        cl.durSeqF   = Math.round(cl.durF * seqFps / cl.srcFps);
        cl.seqStartF = seqHead;
        cl.seqEndF   = seqHead + cl.durSeqF;
        seqHead += cl.durSeqF;
    }
    var totalSeqF = seqHead;

    // ─── build XML ───────────────────────────────────────────────────────────

    var L = [];

    L.push('<?xml version="1.0" encoding="UTF-8"?>');
    L.push('<!DOCTYPE xmeml>');
    L.push('<xmeml version="4">');

    // ── sequence wrapper ──────────────────────────────────────────────────
    L.push('\t<sequence id="sequence-1">');
    L.push('\t\t<duration>' + totalSeqF + '</duration>');
    L.push(rateBlock(seqFps, "\t\t"));
    L.push('\t\t<name>' + xmlEscape(projectBaseName) + '</name>');

    // ── sequence timecode ──────────────────────────────────────────────
    L.push('\t\t<timecode>');
    L.push(rateBlock(seqFps, "\t\t\t"));
    L.push('\t\t\t<string>00:00:00:00</string>');
    L.push('\t\t\t<frame>0</frame>');
    L.push('\t\t\t<displayformat>NDF</displayformat>');
    L.push('\t\t</timecode>');

    // ── media ─────────────────────────────────────────────────────────
    L.push('\t\t<media>');
    L.push('\t\t\t<video>');

    // format block (describes the sequence)
    L.push('\t\t\t\t<format>');
    L.push('\t\t\t\t\t<samplecharacteristics>');
    L.push(rateBlock(seqFps, "\t\t\t\t\t\t"));
    L.push('\t\t\t\t\t\t<width>'  + clips[0].width  + '</width>');
    L.push('\t\t\t\t\t\t<height>' + clips[0].height + '</height>');
    L.push('\t\t\t\t\t\t<anamorphic>FALSE</anamorphic>');
    L.push('\t\t\t\t\t\t<pixelaspectratio>square</pixelaspectratio>');
    L.push('\t\t\t\t\t\t<fielddominance>none</fielddominance>');
    L.push('\t\t\t\t\t</samplecharacteristics>');
    L.push('\t\t\t\t</format>');

    // video track
    L.push('\t\t\t\t<track>');

    for (var i = 0; i < clips.length; i++) {
        var cl = clips[i];
        var url = xmlEscape(toFileURL(cl.path));

        L.push('\t\t\t\t\t<clipitem id="clipitem-' + cl.n + '">');
        L.push('\t\t\t\t\t\t<masterclipid>masterclip-' + cl.n + '</masterclipid>');
        L.push('\t\t\t\t\t\t<name>'    + xmlEscape(cl.name) + '</name>');
        L.push('\t\t\t\t\t\t<enabled>TRUE</enabled>');
        L.push('\t\t\t\t\t\t<duration>' + cl.durF + '</duration>');  // clip duration in srcFps
        L.push(rateBlock(cl.srcFps, "\t\t\t\t\t\t"));               // native source fps
        L.push('\t\t\t\t\t\t<start>' + cl.seqStartF + '</start>');  // sequence position in seqFps
        L.push('\t\t\t\t\t\t<end>'   + cl.seqEndF   + '</end>');
        L.push('\t\t\t\t\t\t<in>'    + cl.srcInF    + '</in>');     // source trim in srcFps
        L.push('\t\t\t\t\t\t<out>'   + cl.srcOutF   + '</out>');
        L.push('\t\t\t\t\t\t<pproTicksIn>'  + pproTicks(cl.srcInF,  cl.srcFps) + '</pproTicksIn>');
        L.push('\t\t\t\t\t\t<pproTicksOut>' + pproTicks(cl.srcOutF, cl.srcFps) + '</pproTicksOut>');
        L.push('\t\t\t\t\t\t<alphatype>none</alphatype>');
        L.push('\t\t\t\t\t\t<pixelaspectratio>square</pixelaspectratio>');
        L.push('\t\t\t\t\t\t<anamorphic>FALSE</anamorphic>');

        // file block
        L.push('\t\t\t\t\t\t<file id="file-' + cl.n + '">');
        L.push('\t\t\t\t\t\t\t<name>'    + xmlEscape(cl.name) + '</name>');
        L.push('\t\t\t\t\t\t\t<pathurl>' + url               + '</pathurl>');
        L.push(rateBlock(cl.srcFps, "\t\t\t\t\t\t\t"));            // file's native fps
        L.push('\t\t\t\t\t\t\t<duration>' + cl.totalDurF + '</duration>'); // in srcFps frames
        L.push('\t\t\t\t\t\t\t<timecode>');
        L.push(rateBlock(cl.srcFps, "\t\t\t\t\t\t\t\t"));
        L.push('\t\t\t\t\t\t\t\t<string>' + framesToTC(cl.fileTCFrame, cl.srcFps) + '</string>');
        L.push('\t\t\t\t\t\t\t\t<frame>'  + cl.fileTCFrame + '</frame>');
        L.push('\t\t\t\t\t\t\t\t<displayformat>NDF</displayformat>');
        L.push('\t\t\t\t\t\t\t</timecode>');
        L.push('\t\t\t\t\t\t\t<media>');
        L.push('\t\t\t\t\t\t\t\t<video>');
        L.push('\t\t\t\t\t\t\t\t\t<samplecharacteristics>');
        L.push(rateBlock(cl.srcFps, "\t\t\t\t\t\t\t\t\t\t")); // match actual file fps
        L.push('\t\t\t\t\t\t\t\t\t\t<width>'  + cl.width  + '</width>');
        L.push('\t\t\t\t\t\t\t\t\t\t<height>' + cl.height + '</height>');
        L.push('\t\t\t\t\t\t\t\t\t\t<anamorphic>FALSE</anamorphic>');
        L.push('\t\t\t\t\t\t\t\t\t\t<pixelaspectratio>square</pixelaspectratio>');
        L.push('\t\t\t\t\t\t\t\t\t\t<fielddominance>none</fielddominance>');
        L.push('\t\t\t\t\t\t\t\t\t</samplecharacteristics>');
        L.push('\t\t\t\t\t\t\t\t</video>');
        L.push('\t\t\t\t\t\t\t\t<audio>');
        L.push('\t\t\t\t\t\t\t\t\t<samplecharacteristics>');
        L.push('\t\t\t\t\t\t\t\t\t\t<depth>16</depth>');
        L.push('\t\t\t\t\t\t\t\t\t\t<samplerate>48000</samplerate>');
        L.push('\t\t\t\t\t\t\t\t\t</samplecharacteristics>');
        L.push('\t\t\t\t\t\t\t\t\t<channelcount>2</channelcount>');
        L.push('\t\t\t\t\t\t\t\t</audio>');
        L.push('\t\t\t\t\t\t\t</media>');
        L.push('\t\t\t\t\t\t</file>');

        // self-link (video)
        L.push('\t\t\t\t\t\t<link>');
        L.push('\t\t\t\t\t\t\t<linkclipref>clipitem-' + cl.n + '</linkclipref>');
        L.push('\t\t\t\t\t\t\t<mediatype>video</mediatype>');
        L.push('\t\t\t\t\t\t\t<trackindex>1</trackindex>');
        L.push('\t\t\t\t\t\t\t<clipindex>' + cl.n + '</clipindex>');
        L.push('\t\t\t\t\t\t</link>');

        // metadata blocks (empty but required by Resolve's parser)
        L.push('\t\t\t\t\t\t<logginginfo>');
        L.push('\t\t\t\t\t\t\t<description></description>');
        L.push('\t\t\t\t\t\t\t<scene></scene>');
        L.push('\t\t\t\t\t\t\t<shottake></shottake>');
        L.push('\t\t\t\t\t\t\t<lognote></lognote>');
        L.push('\t\t\t\t\t\t\t<good></good>');
        L.push('\t\t\t\t\t\t\t<originalvideofilename></originalvideofilename>');
        L.push('\t\t\t\t\t\t\t<originalaudiofilename></originalaudiofilename>');
        L.push('\t\t\t\t\t\t</logginginfo>');
        L.push('\t\t\t\t\t\t<colorinfo>');
        L.push('\t\t\t\t\t\t\t<lut></lut>');
        L.push('\t\t\t\t\t\t\t<lut1></lut1>');
        L.push('\t\t\t\t\t\t\t<asc_sop></asc_sop>');
        L.push('\t\t\t\t\t\t\t<asc_sat></asc_sat>');
        L.push('\t\t\t\t\t\t\t<lut2></lut2>');
        L.push('\t\t\t\t\t\t</colorinfo>');
        L.push('\t\t\t\t\t\t<labels>');
        L.push('\t\t\t\t\t\t\t<label2>Iris</label2>');
        L.push('\t\t\t\t\t\t</labels>');

        L.push('\t\t\t\t\t</clipitem>');
    }

    L.push('\t\t\t\t</track>');
    L.push('\t\t\t</video>');

    // ── audio track (two channels, linked to video clipitems) ─────────────
    L.push('\t\t\t<audio>');
    L.push('\t\t\t\t<track>');
    for (var i = 0; i < clips.length; i++) {
        var cl = clips[i];
        L.push('\t\t\t\t\t<clipitem id="clipitem-' + cl.n + '-audio">');
        L.push('\t\t\t\t\t\t<name>'     + xmlEscape(cl.name) + '</name>');
        L.push('\t\t\t\t\t\t<duration>' + cl.durF            + '</duration>');
        L.push(rateBlock(cl.srcFps, "\t\t\t\t\t\t"));
        L.push('\t\t\t\t\t\t<start>'    + cl.seqStartF       + '</start>');
        L.push('\t\t\t\t\t\t<end>'      + cl.seqEndF         + '</end>');
        L.push('\t\t\t\t\t\t<enabled>TRUE</enabled>');
        L.push('\t\t\t\t\t\t<in>'       + cl.srcInF          + '</in>');
        L.push('\t\t\t\t\t\t<out>'      + cl.srcOutF         + '</out>');
        L.push('\t\t\t\t\t\t<file id="file-' + cl.n + '"/>');
        L.push('\t\t\t\t\t\t<sourcetrack>');
        L.push('\t\t\t\t\t\t\t<mediatype>audio</mediatype>');
        L.push('\t\t\t\t\t\t\t<trackindex>1</trackindex>');
        L.push('\t\t\t\t\t\t</sourcetrack>');
        L.push('\t\t\t\t\t\t<link>');
        L.push('\t\t\t\t\t\t\t<linkclipref>clipitem-' + cl.n + '</linkclipref>');
        L.push('\t\t\t\t\t\t\t<mediatype>video</mediatype>');
        L.push('\t\t\t\t\t\t</link>');
        L.push('\t\t\t\t\t</clipitem>');
    }
    L.push('\t\t\t\t</track>');
    L.push('\t\t\t</audio>');

    L.push('\t\t</media>');
    L.push('\t</sequence>');
    L.push('</xmeml>');

    // ─── save ────────────────────────────────────────────────────────────────

    var _d = new Date();
    var dateStr = String(_d.getFullYear()) + pad2(_d.getMonth() + 1) + pad2(_d.getDate());

    var saveFile;
    if ($.global.__shotRoundtripXMLDir) {
        // Called from roundtrip — save automatically into the shots folder
        var xmlDir = new Folder($.global.__shotRoundtripXMLDir);
        if (!xmlDir.exists) xmlDir.create();
        saveFile = new File(xmlDir.fsName + "/" + projectFileStem + "_shots_" + dateStr + ".xml");
    } else {
        // Standalone — show Save dialog
        var saveDir = proj.file ? proj.file.parent : Folder.desktop;
        saveFile = new File(saveDir.fsName + "/" + projectFileStem + "_shots_" + dateStr + ".xml");
        saveFile = saveFile.saveDlg("Save XML for DaVinci Resolve", "XML files:*.xml,All files:*.*");
        if (!saveFile) return;
    }

    saveFile.encoding = "UTF-8";
    if (!saveFile.open("w")) {
        alert("Failed to write XML file:\n" + saveFile.fsName + "\n\nCheck folder permissions and free disk space.");
        return;
    }
    saveFile.write(L.join("\n"));
    saveFile.close();

})();
