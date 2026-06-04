## SMTV Auto Slides — v2.0.4 (test prerelease)

A small follow-up to v2.0.3 — only the update-nag suppression fix. Shipped as a GitHub prerelease for in-house testing of the deferred-install flow before going public.

### What's fixed

- **Update-required modal no longer re-pops while a v2.0.3 install is staged.** Before the fix, `evaluateUpdateAlert` only suppressed the nag for the current panel session. If the Windows deferred installer was staged but hadn't finished swapping files (e.g. Premiere wasn't fully closed yet), every panel reload re-triggered the modal.
- Now `evaluateUpdateAlert` also reads `update-install-status.json`. If `state` is `staged` or `pending` AND the staged version is greater-than-or-equal-to the live latest, the modal is suppressed — you already engaged with the update flow.
- A LATER, NEWER release (e.g. v2.0.5 dropping while you're staged on v2.0.4) still triggers the nag correctly, because the staged version comparison is against the live latest.

### Notes (technical)

- One additive guard added to `evaluateUpdateAlert` in `client/main.js`. The existing bootstrap behaviour around staged installs is unchanged: status strip still shows the "A staged update is still pending" message.
- No version-bump check needed in the deferred installer; the suppression handles whatever PowerShell writes into the status file.

### Distribution

- Marked as a GitHub **prerelease** so it does NOT appear in `/releases/latest`.
- Only panels with the `smtv-auto-slides-test-updates.flag` flag file in `~/.new-peace-maker/` will see and offer this update. Public users continue to see v2.0.3 as latest.
