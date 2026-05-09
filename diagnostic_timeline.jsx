(function smtvDiagnosticTimeline() {
    var seq = app.project.activeSequence;
    if (!seq) { alert("No active sequence!"); return; }

    // Get in/out points from the sequence
    var inSec  = 0;
    var outSec = 9999;
    try {
        var ip = seq.getInPoint();
        var op = seq.getOutPoint();
        if (ip && typeof ip.seconds === 'number' && ip.seconds >= 0) inSec  = ip.seconds;
        if (op && typeof op.seconds === 'number' && op.seconds  > 0) outSec = op.seconds;
    } catch (e) {}

    var lines = [];
    lines.push("=== SMTV Timeline Diagnostic ===");
    lines.push("Sequence : " + seq.name);
    lines.push("In  : " + inSec.toFixed(3)  + "s  (" + _secToTC(inSec)  + ")");
    lines.push("Out : " + outSec.toFixed(3) + "s  (" + _secToTC(outSec) + ")");
    lines.push("Window: " + (outSec - inSec).toFixed(1) + "s");
    lines.push("");

    var numTracks = seq.videoTracks.length;
    lines.push("Video tracks: " + numTracks);
    lines.push("");

    for (var vi = 0; vi < numTracks; vi++) {
        var track = seq.videoTracks[vi];
        var clips = track.clips;
        var n = clips.numItems;

        if (!n) {
            lines.push("V" + (vi + 1) + ": (empty)");
            continue;
        }

        lines.push("V" + (vi + 1) + ": " + n + " clip(s)");

        for (var ci = 0; ci < n; ci++) {
            var clip    = clips[ci];
            var cStart  = clip.start.seconds;
            var cEnd    = clip.end.seconds;

            // Show clips that overlap the in/out window (±5s margin)
            if (cEnd < inSec - 5 || cStart > outSec + 5) continue;

            var marker = (cStart >= inSec && cStart <= outSec) ? " ◄ IN WINDOW" : "";
            lines.push("  [" + cStart.toFixed(2) + "s → " + cEnd.toFixed(2) + "s]"
                      + "  (" + _secToTC(cStart) + " → " + _secToTC(cEnd) + ")"
                      + "  " + clip.name
                      + marker);
        }
        lines.push("");
    }

    var report = lines.join("\n");
    $.writeln(report);   // also prints to ExtendScript console
    alert(report);

    function _secToTC(s) {
        var fps = 25;
        try { fps = seq.timebase ? parseInt(seq.timebase) : 25; } catch(e){}
        var totalFrames = Math.round(s * fps);
        var fr = totalFrames % fps;
        var ss = Math.floor(totalFrames / fps) % 60;
        var mm = Math.floor(totalFrames / fps / 60) % 60;
        var hh = Math.floor(totalFrames / fps / 3600);
        return _pad(hh) + ":" + _pad(mm) + ":" + _pad(ss) + ":" + _pad(fr);
    }
    function _pad(n) { return n < 10 ? "0" + n : "" + n; }
})();
