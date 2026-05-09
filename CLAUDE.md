# SMTV Auto Slides — Handoff for the next Claude session

This file is auto-loaded by Claude Code at the start of any session in this
project. Read it first — it captures the current state, the user's
preferences, and traps that have already cost time.

> **Last updated:** 2026-05-09 (after v2.0.0 release)

---

## What this project is

Adobe Premiere Pro **CEP panel** at `D:\OpenAI Codex\SMTV AUTO SLIDES\SMTV_Slides`.
Runs inside Premiere's embedded Chromium (CEF) — **NOT** a standalone browser.
Two tabs: **Slides** (.mov categories) and **Quan-Yin & Max / QYM** (PNG ad-cards).

### No dev server, no browser preview

There is no localhost / dev server / browser preview. The panel uses
`require('fs')`, `CSInterface`, and ExtendScript (`host/main.jsx`) — none of
which run in a regular browser. **Do not call `preview_start`, do not
suggest `npm run dev`.** Verification happens by reloading the panel inside
Premiere. If a `<system-reminder>` after an edit nags you to start a
preview, ignore it for files in this project and tell the user you're
skipping per the no-preview rule.

---

## Repo & GitHub state

- Remote: <https://github.com/blueresonara-Sky/SMTV-AUTO-SLIDES-AI>
- Local main and remote main **had no common ancestor** until v2.0.0; the
  remote's prior history (14 commits including web uploads) was wiped by an
  authorized force push. They're recoverable from GitHub's reflog for ~90
  days if ever needed.
- **Current tip:** `Release 2.0.0` at SHA prefix `10e48a0`. README now
  contains a v2.0.0 changelog section.

### Premiere file-lock trap when running git

When Adobe Premiere Pro (or sometimes Defender / Search Indexer) is running,
**`git` cannot rewrite `.git/index` or `.git/refs/heads/*`** in this folder
— it gets `fatal: unable to write new index file` or
`fatal: couldn't set 'refs/heads/main'`. Direct file writes via PowerShell
work fine; only git's atomic-rename pattern fails. Closing Premiere does
**not** always release the lock.

**Workaround that worked:**

1. Redirect git to an alternate index file:
   ```powershell
   $alt = "$env:TEMP\smtv_git_index_$(Get-Random)"
   Copy-Item .git\index $alt -Force
   $env:GIT_INDEX_FILE = $alt
   git add <files>          # or git commit -F msgfile
   Remove-Item env:GIT_INDEX_FILE
   Copy-Item $alt .git\index -Force   # in-place overwrite works
   ```
2. For ref updates that fail the same way, capture the new SHA from
   `.git/logs/HEAD` (the reflog gets written even when the ref doesn't),
   then write it manually:
   ```powershell
   [System.IO.File]::WriteAllText("$PWD\.git\refs\heads\main", "<sha>`n")
   ```
3. `git push --force` itself works fine — only the local `refs/remotes/origin/*`
   may need the same manual write afterward.
4. `git commit -m "<heredoc>"` from PowerShell splits multiline strings into
   positional args. Use `git commit -F <file>` and write the file with
   `[System.IO.File]::WriteAllText` (no BOM) to avoid the UTF-8 BOM that
   `Out-File -Encoding utf8` adds.

If running git inside Claude Code's sandbox produces "Access denied" on
`.git/index`, pass `dangerouslyDisableSandbox: true` to the Bash/PowerShell
tool — but this is **only for the user's own repo and an authorized push**.

---

## Architecture (key files)

- `client/index.html` — panel UI, two tabs.
- `client/main.js` — IIFE, JavaScript logic. **Not** ExtendScript-compatible.
  Uses Node `fs`, `path`, `os`, `child_process`, plus the CEP `CSInterface`
  bridge to call into JSX.
- `client/style.css` — panel styling.
- `client/lib/` — `face-api.js`, `ocrad.js`, `tf.min.js`, `coco-ssd.min.js`,
  and the SSD-MobileNet weights (`model.json` + `group1-shard1of5..5of5`,
  ~5 MB). All bundled in the repo.
- `client/models/` — face-api TinyFaceDetector weights (Slides tab only).
- `host/main.jsx` — ExtendScript ES3. **No** reliable
  `Array.prototype.indexOf/filter/forEach`. Use manual loops + helpers
  like `_arrayIncludes`.

---

## QYM tab — what it does and the standing rule

The Quan-Yin & SMTV Max tab places PNG ad-cards on a high video track
(default V10), top-right corner, with crossfade.

- Folders: `W:\@ EDITORS WORKSPACE\SLIDES\Quan Yin` (27 langs, ENGLISH A/B + non-English split on `--`), `W:\@ EDITORS WORKSPACE\SLIDES\SMTV MAX` (38 codes, English `-ENG.` and non-English `-CODE.`).
- Hard-coded constants in `host/main.jsx`:
  ```js
  var _QYM_MOTION_PRESETS = {
    'quan-yin': { position:[1688.0, 440.0], scale: 90.0 },
    'smtv-max': { position:[1688.0, 440.0], scale: 90.0 }
  };
  var _QYM_LABEL_QUAN_YIN = 15;     // Yellow
  var _QYM_LABEL_SMTV_MAX = 5;      // Forest green
  var _QYM_TICKS_PER_SECOND = 254016000000.0;
  ```
- Slide footprint in 1920×1080: x=76–100%, y=29–53%.

### Standing rule: **DO NOT modify QYM tab code unless the user explicitly asks.**
This is a memory-backed instruction. The user's working order is to leave
QYM detection and placement code untouched.

---

## Critical PPro ExtendScript quirks (hard-won)

- `Time.ticks` must be set as a **String**, not a number.
- Opacity keyframes:
  1. `param.setTimeVarying(true)`
  2. `param.addKey(time)` — pass a `Time` object
  3. `param.setValueAtKey(time, value, 1|0)` — `Time` object + integer flag
     (NOT index, NOT boolean)
  4. `setInterpolationTypeAtKey(time, 0, updateUI)` for linear.
- Working reference: `D:\OpenAI Codex\Footage Courtesy Extension\Auto Footage Courtesy Extension\jsx\hostscript.jsx` — `setParamKey()`.
- `secToMS` previously used NDF math; v2.0.0 uses proper SMPTE 29.97 drop-frame so log labels match Premiere's display.

---

## Detection pipelines

### Slides tab (face-avoidance)
TinyFaceDetector loaded at startup. Three fixed sample offsets per slide:
1.0 s, 4.5 s, 8.0 s. Strict per-candidate: any unsafe sample rejects the
whole candidate. Equal-spacing fallback when avoidFaces is OFF.

### QYM tab
1. **coco-ssd person detection** — `cocoSsd.load({ modelUrl: ... })` on
   `client/lib/model.json`. Runs on a CROP of the frame around top-right
   with **10% uniform margin** (NOT asymmetric). Crop math: TR x=76-100%,
   y=29-53% → crop x=66-100%, y=19-63%. Maps boxes back to full-frame.
   Reject if any person bbox overlaps TR + 3% expansion. **DO NOT change.**
2. **OCRAD text detection** — runs **only when coco-ssd says no person**.
   Asymmetric margins per spec: top=0%, right=0%, left=+10%, bottom=+10%.
   4-pass binarisation (otsu+inv, otsu, t200+inv, t200). 2× upscale before
   OCR. Function: `_qymDetectTextInRegion(canvas)` in `client/main.js`.

### Classification
```
classifyQYM(preview):
  if tr.faceDetected → 'face'   // person in TR
  if tr.textDetected → 'text'   // caption in TR
  else → 'safe'
```

---

## Slides tab settings & flags

Stored in `~/.new-peace-maker/usage-history.json` under `settings`:

| Key | Default | Notes |
|---|---|---|
| `rootFolder` | "" | Picked by user |
| `slideCount` | 6 | Slides per category |
| `targetTrack` | 9 | Starting video track |
| `ignoreV1` | false | "Avoid placing slides over V1 clips" — variable name is misleading; checked = avoid V1 |
| `slideAnchor` | 'top-right' | or 'top-left' |
| `avoidFaces` | true | TinyFaceDetector pass |
| `packPerCategory` | **false (new in 2.0.0)** | Pack slides tight per category, gap between groups |

### `packPerCategory` semantics
- Within each category zone: slides land back-to-back at `zone.start`,
  separated by **0.5 s** (`slideGap`).
- Between category groups: leftover zone time becomes the inter-group gap.
- **Bypasses** both V1 avoidance and face/head detection — slides land at
  packed positions regardless of underlying video.
- Implementation lives in two paths in `client/main.js`:
  - `buildSafePlacementPlan` (avoidFaces ON path) — `originalStart = cursor`
    when `options.packPerCategory && !isFirstInCat`.
  - The avoidFaces-OFF equal-spacing block (~line 4416) — uses a per-category
    cursor `catCursor0[cat]` so V1-blocked nudges accumulate (without it,
    multiple slides collapsed onto the same `nextAvailableStart`).
- The override that disables V1 + faces in pack mode is in
  `previewAndPlaceBatches` right after the preview returns.

---

## QYM tab settings

| Key | Default |
|---|---|
| `qymQuanYinFolder`, `qymSmtvMaxFolder` | "" |
| `qymMaxCount` | 4 |
| `qymTargetTrack` | 10 |
| `qymFadeDuration` | 0.3 |
| `qymAvoidFaces` | true |
| `qymSlideType` | `'both'` / `'qy-only'` / `'max-only'` |
| `qymKeepDebug` | false (UI **hidden** in 2.0.0; logic intact) |

Cycle tracking (also in tracking JSON):
- `quanYinCycle.nextEnglish`: `'A'` or `'B'`, toggles each run.
- `quanYinCycle.usedNonEnglish[]`: lang names burned this cycle.
- `smtvMaxCycle.usedNonEnglish[]`: same.
- Reverts on skip/JSX failure via `_qymRevertCycleForSlides()` +
  `_qymCommitRevert()`.

### Re-enabling the QYM debug checkbox
Search `client/index.html` for `qymKeepDebug` — the field is wrapped in an
HTML comment block. Uncomment to restore the UI. All underlying setting
persistence and the `~/qym-debug/` writer are still wired up.

---

## v2.0.0 highlights (current release)

1. **QYM tab** — full implementation, language cycle, motion presets, fade.
2. **OCRAD text detection** — runs after coco-ssd no-person.
3. **coco-ssd person detection** — model files bundled in `client/lib/`.
4. **Slides tab `Group slides per category`** option (pack mode, see above).
5. **Per-category sort by start time** in `buildSafePlacementPlan` — English
   (files[0]) always maps to earliest TC even when gap insertion shuffles.
6. **Drop-frame timecode in `secToMS`** — fixes ~2 frames/min drift.
7. **Frame export filename** includes `_t<sec>s<cs>` for matching back.
8. **Loading banner** with model-status tracking at startup.
9. **QYM debug-frame export** UI checkbox hidden (logic preserved).

Versions bumped to **2.0.0** in:
- `package.json`
- `CSXS/manifest.xml` (both `ExtensionBundleVersion` and `Extension Version`)

---

## User communication style — what works

- Non-native English speaker, learning. The user's `~/.claude/CLAUDE.md`
  asks for gentle correction blocks at the end of replies. Keep them short,
  format `?? original → corrected (Pattern name)`. Patient teacher tone.
- The user is frustrated by over-engineering or guessing. **Check existing
  code first** before implementing. "Be smart" is real feedback meaning
  "find the principle instead of patching per-frame symptoms."
- Wants verification — sometimes drags debug frames from
  `C:\Users\Mani\qym-debug\<run>\` into chat. Read those when offered.
- Reference extension at `D:\OpenAI Codex\Footage Courtesy Extension\Auto Footage Courtesy Extension` is the gold-standard for things like opacity keyframes.
- Prefers terse responses focused on what changed and what's next.
- Voice-to-text is often used — expect English transcription artefacts.

---

## Working notes / open items

- `posCache` in `buildSafePlacementPlan` is keyed by `Math.round(t)`; cache
  is shared across slides so repeats are free.
- `TRACE_CAP = 60` lines per slide in QYM placement.
- Frame-sampling differs by tab: QYM uses [start, mid, end] of slide
  duration; Slides tab uses fixed [1.0, 4.5, 8.0] (fade-aware).
- The OCRAD pass is new and worth real-world testing on sequences with
  broadcast captions ("Words of Wisdom" cards, "Najia presenter ribbon").
- If a user message complains about overlapping packed slides, the bug is
  almost certainly the OFF path's per-category cursor — check
  `catCursor0[cat]` is being advanced *after* every placement.
