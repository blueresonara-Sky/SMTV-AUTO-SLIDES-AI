## SMTV Auto Slides — v2.0.1

A small UI/UX polish release on top of v2.0.0. No new placement logic.

### What's new

- **Updates moved above the tabs.** The Installed / Latest / Update Now panel now lives at the top of the window, above the *Slides* and *Quan-Yin & Max* tabs, so you can see and trigger updates from either tab without switching back.
- **Compact updates strip.** The Installed and Latest version cards are now shown inline (e.g. `Installed: 2.0.1`, `Latest: 2.0.1`) on a single row, and the whole strip is trimmed down so it doesn't take much vertical space at the top of the panel.
- **Capacity warning on the Slides tab.** Before placement starts, the panel reads your sequence's In/Out range (or the whole sequence if no In/Out is set) and compares it against your requested slide count × number of categories × 9 s. If the slides won't fit, a confirmation dialog tells you:
  - whether the limit is your In/Out range or the whole sequence,
  - how long that window is,
  - how many seconds the request needs and how many seconds short you are,
  - a recommended max per category that does fit.
  You can **Continue anyway** (slides will overlap and burn over each other) or **Cancel** to lower the count first. If the JSX probe can't read a sequence, the check is skipped so nothing blocks you.
- **Bolder, warmer tab labels.** The two tab labels are now bold amber/gold instead of grey, so the tab strip is easier to read at a glance.

### Notes (technical)

- Updates panel moved out of `#panelSlides` in `client/index.html` and into the app shell above `.tab-bar`. New `.update-strip` CSS class controls the compact layout (inline label + version, smaller status box, smaller Update Now button).
- Slides tab `runBtn` handler now does a pre-flight `callJsx('qymGetSequenceWindow', ...)` probe and uses `showUpdateModal(..., { confirm: true })` to ask the user before going ahead when total slide time exceeds the placement window.
- Tab label styling: `.tab-btn` is now bold 14 px in `#f5b942`, active state in `#ffd166` with a matching gold underline.
- No changes to placement math, QYM logic, or detection pipelines.
