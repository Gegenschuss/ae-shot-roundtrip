# diagnostics/

Read-only inspection scripts for debugging the shot-roundtrip pipeline.
None of these mutate the project — they just dump state.

## `dump_timing_state.jsx`

Snapshots every AVLayer's timing state in the active comp and every
nested precomp.  Used to debug time-remap / time-reverse handling on
footage layers and precomps where the roundtrip silently changes the
wrong thing.

### Protocol

Run this *before and after* each roundtrip step you want to inspect,
tagging each snapshot.  Diff the resulting text files to see exactly
what changed.

1. Open the comp.  Make it the active item.
2. `File > Scripts > Run Script File…` → `dump_timing_state.jsx`.
3. Tag the snapshot — e.g. `before`, `after-prep`, `after-confirm`,
   `after-bake`, `after-finish`.
4. Snapshot lands next to the `.aep` as
   `timing_dump_<compName>_<tag>.txt`.
5. Repeat after the next pipeline step.

Two snapshots is the minimum useful pair; for a really nasty bug
(`before` / `after-prep` / `after-bake` / `after-finish`) gives a
per-phase diff.

### What's captured

Per layer (recursively, into every precomp):

- `inPoint` / `outPoint` / `duration` / `startTime` (4-decimal seconds)
- `timeStretch`
- `timeRemapEnabled` + `numKeys` + per-key `(time, value, in/out interp,
  ease handles)`
- Time remap evaluated at `inPoint` and `outPoint` plus a direction
  tag (`ASCENDING` / `DESCENDING` / `FLAT`)
- Source name + source duration
- Layer markers (`cut in` / `cut out` etc.)
- Parent linkage
- Flags: enabled, solo, guide, adjustment

Plus per comp: name, duration, frameRate, frameDuration, comp markers.

Layers with any time effect are tagged `[TIME-EFFECT]` so they're easy
to spot when scrolling.

### Example diff workflow

```
diff timing_dump_mainComp_before.txt timing_dump_mainComp_after-prep.txt
diff timing_dump_mainComp_after-prep.txt timing_dump_mainComp_after-bake.txt
```

Or open both in any text-diff tool (BBEdit, Kaleidoscope, VS Code).

### Sending the dump for help

When asking me to debug, attach both `_before.txt` and `_after.txt`
(or however many phase tags you have).  No screenshots needed —
the text shows exact key times, values, and interpolation modes.
