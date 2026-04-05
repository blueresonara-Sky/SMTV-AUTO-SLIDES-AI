# SMTV Auto Slides - Premiere Pro CEP Extension

SMTV Auto Slides is a Premiere Pro CEP panel that selects slide clips from multiple language folders, imports them into the project, and distributes them across a sequence automatically.

## Categories

The panel works with these category folders:

1. `NEW PEACE MAKER`
2. `Be Vegan Keep Peace`
3. `Forgiveness`
4. `Save the Earth`
5. `Veganism`

## Main Features

1. Choose one root folder that contains all category folders.
2. Pick one English title per category and keep English first.
3. Fill the remaining slide slots with matching non-English versions.
4. Normalize language names so uppercase, lowercase, short codes, and full names are treated as the same language.
5. Track used English titles and used non-English languages locally so the extension rotates through options over time.
6. Support Save the Earth fallback title families such as:
   - `BE KIND`, `BE KIND 2`
   - `BE FRUGAL`, `BE FRUGAL 2`
   - `BE VEG GO GREEN`, `BE VEG GO GREEN2`, `BE VEG GO GREEN 2`
   - related word-order variants such as `GO VEG BE GREEN`
7. Reuse already-used languages for Save the Earth only when necessary, without duplicating a language inside the same Save the Earth batch.
8. Import selected clips into the Premiere project under `Slides`.
9. Place clips across the usable timeline instead of stacking them blindly.
10. Respect sequence In/Out points when they are set.
11. Optionally avoid placing slides during time ranges covered by clips on `V1`.
12. Apply Motion settings to placed slides, except for `Be Vegan Keep Peace`.
13. Let the user choose the slide placement preset:
   - `Top Right`
   - `Top Left`
14. Label English and non-English source items with different Premiere label colors.
15. Check GitHub for updates automatically when the extension starts.
16. Show a green `Update Now` button only when a newer release is available.

## Current Motion Presets

- `Top Right`
  - Position: `1795, 336`
  - Scale: `66`
- `Top Left`
  - Position: `936, 372`
  - Scale: `66`

`Be Vegan Keep Peace` is excluded from the automatic Motion adjustment.

## How Title Selection Works

For each category:

1. The panel scans all titles that have an English version.
2. It prefers English titles that have not been used yet in that category's local history.
3. It prefers titles that can provide more still-unused non-English languages.
4. Randomness is used as a tiebreaker, so results are not globally identical for every user.

History is stored locally per user account:

- Windows: `%USERPROFILE%\.new-peace-maker\usage-history.json`
- macOS: `~/.new-peace-maker/usage-history.json`

This means different users do not automatically share the same title memory unless they are using the same OS account and the same tracking file.

## Update Behavior

Update flow:

1. On startup, the panel checks the latest GitHub Release.
2. If the installed version is current, no update button is shown.
3. If a newer release exists, `Update Now` appears in green.
4. Clicking `Update Now` downloads the release zip, validates it, installs it, and prompts the user to restart Premiere Pro.

For updates to work correctly, each GitHub Release must include a zip asset that contains a valid extension package. The updater can search inside the extracted zip until it finds the extension root.

## Installation

1. Copy the extension folder to your CEP extensions location.
   - Windows: `C:\Program Files (x86)\Common Files\Adobe\CEP\extensions\`
   - macOS: `/Library/Application Support/Adobe/CEP/extensions/`
2. Enable unsigned CEP extensions if needed.
3. Restart Premiere Pro.
4. Open the panel from `Window > Extensions > SMTV Auto Slides`.

## Files

- `CSXS/manifest.xml` - CEP manifest and extension version
- `client/index.html` - panel UI
- `client/style.css` - panel styling
- `client/main.js` - folder scan, title/language selection, local tracking, updater UI and install flow
- `host/main.jsx` - Premiere import, timing logic, V1 avoidance, Motion settings, placement

