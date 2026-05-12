## SMTV Auto Slides — v2.0.2

A startup-speed + responsiveness pass, plus a critical fix for the "track doesn't exist" crash.

### Faster panel boot

- **Deferred model loading.** Detection models (face-api / coco-ssd / OCRAD) no longer load when the panel opens. They only load the first time you click Run with the corresponding *Avoid heads/faces* checkbox ON. Cold-open boot is now under a second instead of 4–5 seconds, and the panel uses zero VRAM until you actually need detection.
- **Scripts moved to `<head>` with `defer`.** The ~4 MB of detection libraries now fetch in parallel while the HTML parses, instead of serialising through the script parser at end-of-body. Big reduction in the "black panel for a few seconds" startup flash.
- **Inline background colour in `<head>`.** CEF no longer paints the WebView area pure black for a frame or two before the stylesheet applies.

### No more frozen panel while editing

- **Toggling an *Avoid heads/faces* checkbox is instant.** Previously, ticking it would kick off a several-second TF.js model load that froze the JS thread, making the rest of the panel unresponsive. Now the toggle just persists the setting. The freeze only happens during the explicit Run step, where waiting is expected.
- **Settings stay editable while models load on Run.** Folder pickers, track numbers, fade duration, slide type — everything stays interactive in the brief window between clicking Run and the placement actually starting. Only the Run button on the tab whose model is loading is locked, with a clear "Loading models…" label.
- **Removed the panel-wide gray-out.** The old `.models-loading` dim that made everything look broken is gone. The loading banner alone tells you what's happening.
- **Loading banner re-shows correctly after the first load.** Previously, once models had loaded once, the banner stayed hidden even when a new model was loading later in the session — you'd see the panel go silent with no explanation. Fixed.
- **No more "All detection models ready" green flash on boot.** Under deferred loading, nothing actually loads on boot, so the false confirmation message is suppressed. The green flash still appears at the real moment models finish loading during a Run.

### Bug fixes

- **`seq.videoTracks.addTrack` crash fixed.** Premiere's ExtendScript API has never had a method to add video tracks programmatically; the old code crashed with `ReferenceError: seq.videoTracks.addTrack is not a function` the moment a user picked a track number higher than what the sequence had. Both the Slides import path and the QYM import path now bail out with a clear instruction instead:
  > Target video track V12 doesn't exist in this sequence (it has only 7 video tracks). Add 5 more in Premiere (right-click the V7 track header → "Add Track…") and try again.
- **Pre-flight track check.** The same error is now detected the moment you click Run, before model loading or scanning, so you don't sit through a long pipeline just to hear "track doesn't exist."
- **Bolder, warmer tab labels** (carry-over polish from the v2.0.1 cycle), more compact Updates strip above the tabs (single line: `Installed: 2.0.2`, `Latest: 2.0.2`), and the capacity warning still runs before placement.

### Notes (technical)

- `loadFaceApiIfNeeded` / `loadCocoSsdIfNeeded` / `loadOcradIfNeeded` are promise-returning wrappers around the underlying init functions. Run handlers `await` them right before the heavy work starts.
- Waiters are pushed onto the model-load queue **before** calling init, to dodge a race when init resolves synchronously (e.g. `initOCRAD`).
- `qymGetSequenceWindow` now also returns `numVideoTracks` so the panel can validate the target track without an extra JSX round-trip.
- The crashing `while (seq.videoTracks.numTracks <= trackIndex) { seq.videoTracks.addTrack(); }` blocks in `host/main.jsx` were replaced with explicit early-return error JSON in both placement entry points.
- The QYM Run click listener is now a thin pre-flight wrapper: it does the track check, loads models if needed, then hands off to the existing `runQuanYinAndMax` unchanged. The QYM placement logic itself is untouched.
