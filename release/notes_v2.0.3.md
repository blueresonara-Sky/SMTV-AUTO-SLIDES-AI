## SMTV Auto Slides — v2.0.3

A safety pass on placement pre-flight, a robustness fix for In/Out point handling, and a much more visible update prompt.

### Don't burn slides on top of existing clips

- **Target track must be empty.** Both tabs now check the chosen video track before doing any work — if any clip already overlaps the placement range (In/Out marks if set, otherwise the full sequence), the run is cancelled before models load or frames are exported. Pick an empty track or clear the existing clips, then click Run again.
- **Track-existence check covers the empty check too.** This runs in the same pre-flight wrapper as the v2.0.2 "track doesn't exist" check, so both checks fail fast (under a second after Run is clicked) instead of after a long pipeline.
- **Implementation:** `qymGetSequenceWindow` was extended with an optional `targetTrack` payload. When provided, it scans that track and returns `targetTrackClipCount` — the number of clips overlapping the placement window. The QYM Run handler and the Slides Run handler both pass the user's chosen track and bail out if the count is > 0.

### Loud, unmistakable pre-flight errors

- **Red status banner.** When a pre-flight check cancels a run, the status box turns red, gets a 2-pixel red border + glow ring, and the message is prefixed with **⛔**. Stays in that state until the next normal status update.
- **Red modal that pops automatically.** Same message also appears in a centered modal with a red header gradient and a red OK button. Backdrop click doesn't dismiss it — the user has to click OK, so it can't be missed by someone glancing away.
- **CSS-scoped variant.** The red modal style is gated behind a new `.modal-backdrop.error-alert` class. Other modals (install-confirm, capacity warning, no-In/Out confirm) keep their compact default look.

### Hard-to-miss update prompt when a new version is available

- **Big green modal auto-pops.** After the live GitHub check returns with a confirmed newer release, a modal opens automatically with the release notes. Bigger card (640px wide), bold 26-pixel title on a green gradient header, and a chunky **Update Now** primary button. The "Later" button stays modest so it doesn't compete.
- **Pops at most once per panel session.** If you dismiss it with "Later", the panel stays fully usable and the modal won't keep coming back until you reopen the panel.
- **No nag if offline.** If the live check fails (network down, GitHub unreachable), the modal stays silent — we only prompt when we've actually confirmed a downloadable update exists.
- **Stays fully usable.** Earlier experiments with a hard lock were dropped in favour of a clean, in-your-face modal. Nothing in the panel gets disabled.
- **Single Update Now button.** Clicking Update Now in the alert modal goes straight into the install pipeline — no redundant second confirmation modal.

### In/Out point safety: don't trust garbage from Premiere

- **`_getPlacementWindow` now clamps `in`/`out` into `[0, usedLen]`.** Premiere sequences usually have a start timecode of `01:00:00:00` rather than `00:00:00:00`. If the user dropped an In point at displayed `00:00:00:00` (one hour BEFORE the sequence's displayed start), Premiere returned a sentinel/negative seconds value like `−400 000 s`. The old code trusted it and produced a window like `Window: −6666:40:00 – 5:14:12` — about 4.6 days wide — so capacity checks passed trivially, the AD/Slogan overlap math didn't subtract anything, and slides ended up in nonsense positions.
- **New behaviour:** if In comes back negative, it's snapped to sequence start (`0 s`) and the Out point is preserved. If Out is past the end of the sequence, it's clamped to `usedLen`. If both clamped values still satisfy `out > in`, the In/Out window is used as-is. Otherwise we fall back to the full sequence — same as before.
- **Helps both tabs.** `_getPlacementWindow` is shared infrastructure used by both the Slides and the QYM tab.

### Notes (technical)

- `host/main.jsx`:
  - `_getPlacementWindow` no longer blindly trusts `seq.getInPointAsTime()` / `seq.getOutPointAsTime()`. It clamps each endpoint into `[0, usedLen + 0.5 s]` before deciding whether to use them.
  - `qymGetSequenceWindow` accepts an optional `{ targetTrack: <1-based number> }` payload. When provided and valid, scans that track for clips whose interval overlaps the placement window (with a half-millisecond epsilon to ignore touching edges) and returns `targetTrackClipCount` in the response.
- `client/main.js`:
  - New `showPreflightError(tab, message)` helper that paints the relevant status box red, prefixes the message with ⛔, and pops a red `.error-alert` modal with title "Cannot run".
  - `setStatus` / `setQymStatus` auto-clear the `.is-error` class on the next normal update.
  - `installLatestUpdate` was split: the actual install pipeline (download → extract → stage installer) was extracted into a `_performUpdateInstall` helper. The "Update available" alert modal calls `_performUpdateInstall` directly, skipping the redundant second confirmation modal.
  - `setModalOpen` switched from clobbering `className` to using `classList.add/remove('is-open')`, so variant classes (`update-alert`, `error-alert`) survive open/close.
  - `evaluateUpdateAlert` gates the modal auto-pop on a successful live check + a downloadable release in hand + a once-per-session `updateAlertShown` flag.
- `client/style.css`:
  - `.modal-backdrop.update-alert` — bigger green-headed variant for the update prompt.
  - `.modal-backdrop.error-alert` — red-headed variant for blocking pre-flight errors.
  - `#status.is-error` / `#qymStatus.is-error` — red banner state for the status boxes.
