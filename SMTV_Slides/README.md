# SMTV Auto Slides - Premiere Pro CEP Extension

This extension helps place slide clips across a Premiere Pro sequence for these categories:

1. `NEW PEACE MAKER`
2. `Be Vegan Keep Peace`
3. `Forgiveness`
4. `Save the Earth`
5. `Veganism`

## What It Does

1. Lets you choose a root folder that contains the category folders.
2. Scans language folders and slide files.
3. Picks one English title per category and keeps English first.
4. Fills the remaining slide slots with matching non-English versions.
5. Normalizes language names so short codes, full names, and capitalization are treated as the same language.
6. Tracks used English titles and used non-English languages in:
   - Windows: `%USERPROFILE%\.new-peace-maker\usage-history.json`
   - macOS: `~/.new-peace-maker/usage-history.json`
7. Supports Save the Earth title-family fallback for variants such as:
   - `BE KIND`, `BE KIND 2`
   - `BE FRUGAL`, `BE FRUGAL 2`
   - `BE VEG GO GREEN`, `BE VEG GO GREEN2`, `BE VEG GO GREEN 2`
   - similar word-order variants such as `GO VEG BE GREEN`
8. Imports the selected clips into the Premiere project under `Slides`.
9. Places all selected clips on the chosen video track using one shared interval.
10. Can avoid placing slides during time ranges occupied by clips on `V1`.
11. Uses the sequence In/Out range when present; otherwise it uses the full usable sequence length.
12. Applies Motion scale and position to placed slides, except for `Be Vegan Keep Peace`.
13. Labels English and non-English slide source items with different label colors.
14. Can check GitHub Releases for updates and install a newer zip package after confirmation.

## Important Notes

- The first placed clip in each category is always the English file.
- If a chosen title does not have enough matching languages, the extension uses the maximum available and warns you.
- Save the Earth can reuse already-used languages from other categories when needed, without duplicating languages within the same batch.
- The folder picker uses CEP/Chromium directory selection support.
- Update checking requires a GitHub repo in `owner/repo` format and a release zip that contains `CSXS/manifest.xml`.

## Install

1. Copy the extension folder to your CEP extensions location.
   - Windows: `C:\Program Files (x86)\Common Files\Adobe\CEP\extensions\`
   - macOS: `/Library/Application Support/Adobe/CEP/extensions/`
2. Enable unsigned CEP extensions if needed.
3. Restart Premiere Pro.
4. Open the panel from `Window > Extensions > SMTV Auto Slides`.

## Files

- `CSXS/manifest.xml` - CEP manifest
- `client/index.html` - panel UI
- `client/style.css` - panel styling
- `client/main.js` - scan folders, choose title/languages, save history, check/install updates, call JSX
- `host/main.jsx` - import clips, calculate placement timing, avoid blocked V1 ranges, set motion, place clips

## Release Packaging

When creating a release zip, the package must contain the extension root with:

- `CSXS/manifest.xml`
- `client/`
- `host/`
- `README.md`

The updater searches inside the extracted zip until it finds a valid extension root.
