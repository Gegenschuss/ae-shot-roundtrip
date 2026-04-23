/*
 *       _____                          __
 *      / ___/__ ___ ____ ___  ___ ____/ /  __ _____ ___
 *     / (_ / -_) _ `/ -_) _ \(_-</ __/ _ \/ // (_-<(_-<
 *     \___/\__/\_, /\__/_//_/___/\__/_//_/\_,_/___/___/
 *             /___/
 */

/*
================================================================================
  WRITE README SCAFFOLDS — shared helper
  After Effects ExtendScript
================================================================================

Single source of truth for the two README.txt files that the Roundtrip/
tree scaffolds automatically:

  - Roundtrip/README.txt           (handoff tree overview, do-not-rename)
  - Roundtrip/_grade/README.txt    (Resolve delivery preset, naming rules)

Used by both shot-roundtrip/shot_roundtrip.jsx (primary creator) and
import-renders/import_renders.jsx (safety net, in case the user runs
import before a full roundtrip).

Exposes two globals after `$.evalFile`:

  writeRoundtripReadme(folder)   -> true if written, false if skipped
  writeGradeReadme(folder)       -> true if written, false if skipped

Both are write-once: an existing README.txt is never overwritten.
================================================================================
*/

function writeRoundtripReadme(folder) {
    if (!folder || !folder.exists) return false;
    var readme = new File(folder.fsName + "/README.txt");
    if (readme.exists) return false;
    readme.encoding = "UTF-8";
    if (!readme.open("w")) return false;
    readme.write(
        "Roundtrip - VFX shot handoff folder\n" +
        "====================================\n" +
        "\n" +
        "Created by the Gegenschuss AE Shot Roundtrip \"Shot Roundtrip\" script.\n" +
        "This is the canonical location for everything that leaves\n" +
        "After Effects during a project: plates, VFX returns, grades,\n" +
        "and Dynamic Link wrappers.\n" +
        "\n" +
        "Structure\n" +
        "---------\n" +
        "  Roundtrip/\n" +
        "    {prefix}_010/\n" +
        "      plate/            <- rendered plate (.mov)\n" +
        "      render/           <- VFX return from Nuke etc.\n" +
        "    {prefix}_020/\n" +
        "      plate/\n" +
        "      render/\n" +
        "    _grade/             <- Resolve graded returns (flat, shared)\n" +
        "    dynamicLink/        <- Dynamic Link wrapper comps\n" +
        "    {project}_Comp.nk   <- AppendClip master Nuke script (optional)\n" +
        "    {project}.xml       <- Premiere FCPXML (optional)\n" +
        "\n" +
        "IMPORTANT - DO NOT RENAME anything inside this tree.\n" +
        "----------------------------------------------------\n" +
        "The AE scripts rely on a strict naming contract:\n" +
        "\n" +
        "  AE comps  {prefix}_NNN_comp   <->   disk  {prefix}_NNN/\n" +
        "  VFX renders live in               {shot}/render/\n" +
        "  Resolve grades live flat in       _grade/\n" +
        "  Grades match comps by filename prefix (KM_010_* -> KM_010_comp)\n" +
        "\n" +
        "Renaming any of these silently breaks the roundtrip. If a shot\n" +
        "is dropped, just delete its folder. If a new shot is added,\n" +
        "re-run Shot Roundtrip and it will slot in using the next free\n" +
        "number in the increment-of-10 scheme.\n" +
        "\n" +
        "Source & documentation\n" +
        "----------------------\n" +
        "  https://github.com/Gegenschuss/ae-shot-roundtrip\n" +
        "\n" +
        "Built by Gegenschuss  -  https://gegenschuss.com\n"
    );
    readme.close();
    return true;
}

function writeGradeReadme(folder) {
    if (!folder || !folder.exists) return false;
    var readme = new File(folder.fsName + "/README.txt");
    if (readme.exists) return false;
    readme.encoding = "UTF-8";
    if (!readme.open("w")) return false;
    readme.write(
        "Resolve grades - how to use this folder\n" +
        "=========================================\n" +
        "\n" +
        "Drop DaVinci Resolve-graded clips here. The AE \"Import Returns\"\n" +
        "script scans this folder and matches files back to their shot\n" +
        "comps by filename prefix.\n" +
        "\n" +
        "Naming convention\n" +
        "-----------------\n" +
        "Each file must START with the shot name, followed by a\n" +
        "non-alphanumeric character:\n" +
        "\n" +
        "  KM_010_grade_v01.mov   ->  matches comp KM_010_comp\n" +
        "  KM_010_grade_v02.mov   ->  newer version, stacks on top\n" +
        "  KM_020_grade_v01.mov   ->  matches comp KM_020_comp\n" +
        "\n" +
        "The newest version per shot wins. Older versions are imported\n" +
        "but disabled so you can A/B them in AE.\n" +
        "\n" +
        "Resolve Deliver preset\n" +
        "----------------------\n" +
        "  Render         Individual clips\n" +
        "  Format         QuickTime (match your VFX render codec,\n" +
        "                            e.g. ProRes 422 HQ)\n" +
        "  Resolution     Same as source\n" +
        "  Destination    this folder\n" +
        "  Filename       Custom\n" +
        "  Pattern        %Source Name_grade_v01\n" +
        "\n" +
        "Bump the version (_v01 -> _v02) in the preset before every new\n" +
        "delivery pass. Never overwrite -- older looks should stay\n" +
        "recoverable.\n" +
        "\n" +
        "IMPORTANT - DO NOT RENAME this folder.\n" +
        "--------------------------------------\n" +
        "The name \"_grade\" is hard-coded in the Import Returns script.\n" +
        "Renaming it silently breaks grade import.\n" +
        "\n" +
        "Safe to delete the folder entirely if you don't use Resolve\n" +
        "grades -- the script will recreate it on the next run.\n" +
        "\n" +
        "Source: https://github.com/Gegenschuss/ae-shot-roundtrip\n"
    );
    readme.close();
    return true;
}
