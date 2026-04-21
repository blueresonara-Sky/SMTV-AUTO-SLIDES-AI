function _npmJson(obj) {
    if (typeof JSON !== 'undefined' && JSON.stringify) {
        return JSON.stringify(obj);
    }
    return obj.toSource();
}

function _npmParse(jsonStr) {
    if (typeof JSON !== 'undefined' && JSON.parse) {
        return JSON.parse(jsonStr);
    }
    return eval('(' + jsonStr + ')');
}

function _findChildBinByName(parentItem, name) {
    if (!parentItem || !parentItem.children) return null;
    for (var i = 0; i < parentItem.children.numItems; i++) {
        var child = parentItem.children[i];
        if (child && child.type === ProjectItemType.BIN && child.name === name) {
            return child;
        }
    }
    return null;
}

function _findOrCreateBin(parentItem, name) {
    var existing = _findChildBinByName(parentItem, name);
    if (existing) return existing;
    return parentItem.createBin(name);
}

function _findProjectItemByMediaPath(binItem, mediaPath) {
    if (!binItem || !binItem.children) return null;
    for (var i = 0; i < binItem.children.numItems; i++) {
        var child = binItem.children[i];
        if (child) {
            if (child.type === ProjectItemType.BIN) {
                var nested = _findProjectItemByMediaPath(child, mediaPath);
                if (nested) return nested;
            } else {
                try {
                    if (child.getMediaPath && child.getMediaPath() === mediaPath) {
                        return child;
                    }
                } catch (e) {}
            }
        }
    }
    return null;
}

function _findProjectItemByMediaPathAnywhere(parentItem, mediaPath) {
    if (!parentItem) return null;
    if (parentItem.type !== ProjectItemType.BIN) {
        try {
            if (parentItem.getMediaPath && parentItem.getMediaPath() === mediaPath) {
                return parentItem;
            }
        } catch (e) {}
        return null;
    }
    return _findProjectItemByMediaPath(parentItem, mediaPath);
}

function _getSequenceMaxEndSeconds(seq) {
    var maxSec = 0;
    if (!seq) return maxSec;

    var vtCount = seq.videoTracks.numTracks;
    for (var v = 0; v < vtCount; v++) {
        var vTrack = seq.videoTracks[v];
        for (var vc = 0; vc < vTrack.clips.numItems; vc++) {
            var vClip = vTrack.clips[vc];
            if (vClip && vClip.end && vClip.end.seconds > maxSec) {
                maxSec = vClip.end.seconds;
            }
        }
    }

    var atCount = seq.audioTracks.numTracks;
    for (var a = 0; a < atCount; a++) {
        var aTrack = seq.audioTracks[a];
        for (var ac = 0; ac < aTrack.clips.numItems; ac++) {
            var aClip = aTrack.clips[ac];
            if (aClip && aClip.end && aClip.end.seconds > maxSec) {
                maxSec = aClip.end.seconds;
            }
        }
    }

    return maxSec;
}

function _ensureProjectItemInBin(binItem, mediaPath) {
    if (!binItem || !mediaPath) return null;

    var existing = _findProjectItemByMediaPath(binItem, mediaPath);
    if (existing) return existing;

    var rootExisting = _findProjectItemByMediaPathAnywhere(app.project.rootItem, mediaPath);
    if (rootExisting) {
        try {
            rootExisting.moveBin(binItem);
        } catch (moveErr) {}
        existing = _findProjectItemByMediaPath(binItem, mediaPath);
        if (existing) return existing;
        return rootExisting;
    }

    try {
        app.project.importFiles([mediaPath], false, binItem, false);
    } catch (e) {
        return null;
    }

    return _findProjectItemByMediaPath(binItem, mediaPath);
}

function _getProjectItemDurationSeconds(projectItem) {
    if (!projectItem) return 0;
    try {
        var inPoint = projectItem.getInPoint ? projectItem.getInPoint() : null;
        var outPoint = projectItem.getOutPoint ? projectItem.getOutPoint(1) : null;
        var inSeconds = inPoint && typeof inPoint.seconds === 'number' ? inPoint.seconds : 0;
        var outSeconds = outPoint && typeof outPoint.seconds === 'number' ? outPoint.seconds : 0;
        if (outSeconds > inSeconds) {
            return outSeconds - inSeconds;
        }
    } catch (e) {}
    return 0;
}

function _getTrackOccupiedRanges(track) {
    var ranges = [];
    if (!track || !track.clips) return ranges;

    for (var i = 0; i < track.clips.numItems; i++) {
        var clip = track.clips[i];
        if (!clip || !clip.start || !clip.end) continue;
        ranges.push({
            start: clip.start.seconds,
            end: clip.end.seconds
        });
    }

    ranges.sort(function (a, b) {
        return a.start - b.start;
    });

    return ranges;
}

function _clipRangesToWindow(ranges, windowStart, windowEnd) {
    var clipped = [];
    if (!ranges || !ranges.length) return clipped;

    for (var i = 0; i < ranges.length; i++) {
        var range = ranges[i];
        var start = Math.max(windowStart, range.start);
        var end = Math.min(windowEnd, range.end);
        if (end > start) {
            clipped.push({
                start: start,
                end: end
            });
        }
    }

    return clipped;
}

function _mergeRanges(ranges) {
    if (!ranges || !ranges.length) return [];

    var sorted = ranges.slice().sort(function (a, b) {
        return a.start - b.start;
    });
    var merged = [sorted[0]];

    for (var i = 1; i < sorted.length; i++) {
        var current = sorted[i];
        var last = merged[merged.length - 1];
        if (current.start <= last.end) {
            if (current.end > last.end) {
                last.end = current.end;
            }
        } else {
            merged.push({ start: current.start, end: current.end });
        }
    }

    return merged;
}

function _getAvailableRanges(totalLengthSeconds, blockedRanges) {
    var available = [];
    var cursor = 0;
    var mergedBlocked = _mergeRanges(blockedRanges);

    for (var i = 0; i < mergedBlocked.length; i++) {
        var blocked = mergedBlocked[i];
        var start = Math.max(0, blocked.start);
        var end = Math.min(totalLengthSeconds, blocked.end);
        if (start > cursor) {
            available.push({ start: cursor, end: start });
        }
        if (end > cursor) {
            cursor = end;
        }
    }

    if (cursor < totalLengthSeconds) {
        available.push({ start: cursor, end: totalLengthSeconds });
    }

    var cleaned = [];
    for (var j = 0; j < available.length; j++) {
        if (available[j].end > available[j].start) {
            cleaned.push(available[j]);
        }
    }

    return cleaned;
}

function _getTotalRangeLength(ranges) {
    var total = 0;
    if (!ranges) return total;
    for (var i = 0; i < ranges.length; i++) {
        total += Math.max(0, ranges[i].end - ranges[i].start);
    }
    return total;
}

function _mapCompressedTimeToSequenceTime(compressedSeconds, availableRanges) {
    if (!availableRanges || !availableRanges.length) {
        return compressedSeconds;
    }

    var remaining = compressedSeconds;
    for (var i = 0; i < availableRanges.length; i++) {
        var range = availableRanges[i];
        var length = range.end - range.start;
        if (remaining <= length) {
            return range.start + remaining;
        }
        remaining -= length;
    }

    return availableRanges[availableRanges.length - 1].end;
}

function _getPlacementWindow(seq, fallbackEndSeconds) {
    var start = 0;
    var end = fallbackEndSeconds;

    try {
        if (seq && seq.getInPointAsTime && seq.getOutPointAsTime) {
            var inTime = seq.getInPointAsTime();
            var outTime = seq.getOutPointAsTime();
            var inSeconds = inTime && typeof inTime.seconds === 'number' ? inTime.seconds : 0;
            var outSeconds = outTime && typeof outTime.seconds === 'number' ? outTime.seconds : 0;

            if (outSeconds > inSeconds) {
                start = inSeconds;
                end = outSeconds;
            }
        }
    } catch (e) {}

    if (end <= start) {
        start = 0;
        end = fallbackEndSeconds;
    }

    return {
        start: start,
        end: end
    };
}

function _joinPath(dirPath, leafName) {
    if (!dirPath) return leafName || '';
    if (!leafName) return dirPath;
    var sep = '/';
    if (Folder.fs !== 'Macintosh') {
        sep = '\\';
    }
    if (dirPath.charAt(dirPath.length - 1) === '\\' || dirPath.charAt(dirPath.length - 1) === '/') {
        return dirPath + leafName;
    }
    return dirPath + sep + leafName;
}

function _ensureFolder(folderPath) {
    if (!folderPath) return false;
    var folder = new Folder(folderPath);
    if (folder.exists) return true;
    return folder.create();
}

function _sanitizeFileStem(name) {
    return String(name || 'frame').replace(/[^\w.-]+/g, '_');
}

function _findExistingFramePath(basePathWithoutExtension) {
    var candidates = [
        basePathWithoutExtension,
        basePathWithoutExtension + '.png',
        basePathWithoutExtension + '.png.png'
    ];

    for (var i = 0; i < candidates.length; i++) {
        var file = new File(candidates[i]);
        if (file.exists) {
            return file.fsName;
        }
    }

    return '';
}

function _waitForFile(stemPath, maxWaitMs) {
    // Poll for the file since exportFramePNG writes asynchronously.
    var candidates = [
        stemPath,
        stemPath + '.png',
        stemPath + '.png.png'
    ];
    var waited = 0;
    var pollMs = 120;
    while (waited <= maxWaitMs) {
        for (var i = 0; i < candidates.length; i++) {
            var f = new File(candidates[i]);
            if (f.exists && f.length > 0) return f.fsName;
        }
        // $.sleep is synchronous in ExtendScript
        try { $.sleep(pollMs); } catch (e) {}
        waited += pollMs;
    }
    return '';
}

function _exportSequenceFrameAtSeconds(seq, seconds, outputStemPath) {
    var result = { ok: false, path: '', error: '' };
    if (!seq || !outputStemPath) {
        result.error = 'Missing sequence or output path.';
        return result;
    }

    var originalPosition = null;
    try { originalPosition = seq.getPlayerPosition(); } catch (e) {}

    try {
        var exportTime = new Time();
        exportTime.seconds = Math.max(0, seconds || 0);
        seq.setPlayerPosition(exportTime.ticks);

        // Small pause so Premiere registers the new CTI position
        try { $.sleep(80); } catch (e) {}

        app.enableQE();
        var qeSequence = qe && qe.project ? qe.project.getActiveSequence() : null;
        if (!qeSequence || typeof qeSequence.exportFramePNG !== 'function' || !qeSequence.CTI) {
            result.error = 'QE exportFramePNG is unavailable.';
            return result;
        }

        var timecode = qeSequence.CTI.timecode;
        if (!timecode) {
            result.error = 'Could not read CTI timecode.';
            return result;
        }

        // Delete any pre-existing file at the path to avoid stale hits
        var candidates = [outputStemPath, outputStemPath + '.png', outputStemPath + '.png.png'];
        for (var d = 0; d < candidates.length; d++) {
            try { var df = new File(candidates[d]); if (df.exists) df.remove(); } catch (e) {}
        }

        qeSequence.exportFramePNG(timecode, outputStemPath);

        // Wait up to 4 seconds for the async file write to complete
        result.path = _waitForFile(outputStemPath, 4000);
        result.ok = !!result.path;
        if (!result.ok) {
            result.error = 'exportFramePNG ran but no file appeared within 4s. Timecode: ' + timecode;
        }
    } catch (err) {
        result.error = err.toString();
    } finally {
        try {
            if (originalPosition && originalPosition.ticks) {
                seq.setPlayerPosition(originalPosition.ticks);
            }
        } catch (restoreErr) {}
    }

    return result;
}

function diagnosTimelineClips() {
    var seq = app.project.activeSequence;
    if (!seq) return JSON.stringify({ error: 'No active sequence' });

    var inSec  = 0;
    var outSec = 9999;
    try {
        var ip = seq.getInPoint();
        var op = seq.getOutPoint();
        if (ip && typeof ip.seconds === 'number' && ip.seconds >= 0) inSec  = ip.seconds;
        if (op && typeof op.seconds === 'number' && op.seconds  > 0) outSec = op.seconds;
    } catch (e) {}

    var fps = 25;
    try { if (seq.timebase) fps = Math.round(1 / parseFloat(seq.timebase)); } catch(e) {}

    function secToTC(s) {
        var f  = Math.round(s * fps);
        var fr = f % fps;
        var ss = Math.floor(f / fps) % 60;
        var mm = Math.floor(f / fps / 60) % 60;
        var hh = Math.floor(f / fps / 3600);
        function p(n) { return n < 10 ? '0'+n : ''+n; }
        return p(hh)+':'+p(mm)+':'+p(ss)+':'+p(fr);
    }

    var lines = [];
    lines.push('Sequence: ' + seq.name);
    lines.push('In:  ' + inSec.toFixed(2) + 's  (' + secToTC(inSec)  + ')');
    lines.push('Out: ' + outSec.toFixed(2) + 's  (' + secToTC(outSec) + ')');
    lines.push('');

    for (var vi = 0; vi < seq.videoTracks.length; vi++) {
        var track = seq.videoTracks[vi];
        var clips = track.clips;
        var n     = clips.numItems;
        if (!n) { lines.push('V'+(vi+1)+': (empty)'); continue; }
        lines.push('V'+(vi+1)+': '+n+' clip(s)');
        for (var ci = 0; ci < n; ci++) {
            var clip   = clips[ci];
            var cStart = clip.start.seconds;
            var cEnd   = clip.end.seconds;
            if (cEnd < inSec - 5 || cStart > outSec + 5) continue;
            var flag = (cStart >= inSec - 0.1 && cStart <= outSec + 0.1) ? ' ◄' : '';
            lines.push('  ['+cStart.toFixed(2)+'s → '+cEnd.toFixed(2)+'s]'
                      +'  '+secToTC(cStart)+' → '+secToTC(cEnd)
                      +'  '+clip.name+flag);
        }
        lines.push('');
    }

    return lines.join('\n');
}

function _exportPlacementSampleFrames(seq, startSeconds, clipDurationSeconds, analysisDir, fileStemBase) {
    var framePaths = [];
    var sampleTimes = [];
    var exportErrors = [];

    var slideDuration = 9.0;  // SMTV slides are always 9 seconds
    var fadeIn        = 1.0;

    // Fixed samples spread across the slide window.
    // 8.0s is included specifically to catch title text that fades in ~1-2s
    // after a scene cut — the 6s sample often lands only 0.3s into a new clip
    // where text is still semi-transparent and unreadable by OCRAD.
    var sampleOffsets = [
        fadeIn,                   // 1.0s — after fade-in
        slideDuration / 2,        // 4.5s — mid-point
        slideDuration - fadeIn    // 8.0s — near end, catches late-appearing text
    ];

    for (var s = 0; s < sampleOffsets.length; s++) {
        var sampleSeconds = startSeconds + sampleOffsets[s];
        sampleTimes.push(sampleSeconds);
        var exportResult = _exportSequenceFrameAtSeconds(seq, sampleSeconds, _joinPath(analysisDir, fileStemBase + '_sample_' + (s + 1)));
        framePaths.push(exportResult.path || '');
        if (!exportResult.ok && exportResult.error) {
            exportErrors.push('sample ' + (s + 1) + ': ' + exportResult.error);
        }
    }

    return {
        framePaths: framePaths,
        sampleTimes: sampleTimes,
        exportErrors: exportErrors
    };
}

var _NPM_LABEL_ENGLISH = 4;
var _NPM_LABEL_NON_ENGLISH = 8;
var _NPM_MOTION_PRESETS = {
    'top-right': {
        position: [1795.0, 336.0],
        scale: 66.0
    },
    'top-left': {
        position: [936.0, 372.0],
        scale: 66.0
    }
};

function _findTrackItemByStartSeconds(track, startSeconds) {
    if (!track || !track.clips) return null;
    var tolerance = 0.05;
    for (var i = 0; i < track.clips.numItems; i++) {
        var clip = track.clips[i];
        if (!clip || !clip.start) continue;
        if (Math.abs(clip.start.seconds - startSeconds) <= tolerance) {
            return clip;
        }
    }
    return null;
}

function _setProjectItemLabel(projectItem, isEnglish) {
    if (!projectItem || typeof projectItem.setColorLabel !== 'function') return;
    try {
        projectItem.setColorLabel(isEnglish ? _NPM_LABEL_ENGLISH : _NPM_LABEL_NON_ENGLISH);
    } catch (e) {}
}

function _findComponentByMatchName(trackItem, matchName, displayName) {
    if (!trackItem || !trackItem.components) return null;
    for (var i = 0; i < trackItem.components.numItems; i++) {
        var component = trackItem.components[i];
        if (!component) continue;
        if (matchName && component.matchName === matchName) return component;
        if (displayName && component.displayName === displayName) return component;
    }
    return null;
}

function _findComponentProperty(component, matchName, displayName) {
    if (!component || !component.properties) return null;
    for (var i = 0; i < component.properties.numItems; i++) {
        var property = component.properties[i];
        if (!property) continue;
        if (matchName && property.matchName === matchName) return property;
        if (displayName && property.displayName === displayName) return property;
    }
    return null;
}

function _setComponentPropertyValue(property, value) {
    if (!property || typeof property.setValue !== 'function') return;
    try {
        if (typeof property.setTimeVarying === 'function') {
            property.setTimeVarying(false);
        }
    } catch (e) {}
    try {
        property.setValue(value, 1);
    } catch (e) {}
}

function _getSequenceRelativePosition(seq, xy) {
    var width = 1920;
    var height = 1080;

    try {
        if (seq && seq.videoFrameWidth) {
            width = parseFloat(seq.videoFrameWidth) || width;
        }
        if (seq && seq.videoFrameHeight) {
            height = parseFloat(seq.videoFrameHeight) || height;
        }
    } catch (e) {}

    return [
        xy[0] / width,
        xy[1] / height
    ];
}

function _getMotionPreset(anchor) {
    var key = anchor || 'top-right';
    return _NPM_MOTION_PRESETS[key] || _NPM_MOTION_PRESETS['top-right'];
}

function _setMotionPosition(property, xy, seq) {
    if (!property || typeof property.setValue !== 'function' || !xy || xy.length < 2) return;
    var relativeXY = _getSequenceRelativePosition(seq, xy);

    try {
        if (typeof property.setTimeVarying === 'function') {
            property.setTimeVarying(false);
        }
    } catch (e) {}

    try {
        var currentValue = property.getValue();
        if (currentValue && typeof currentValue.length !== 'undefined' && currentValue.length >= 2) {
            currentValue[0] = relativeXY[0];
            currentValue[1] = relativeXY[1];
            property.setValue(currentValue, 1);
            return;
        }
    } catch (e) {}

    try {
        property.setValue([relativeXY[0], relativeXY[1]], 1);
    } catch (e) {}

    try {
        property.setValue([parseFloat(relativeXY[0]), parseFloat(relativeXY[1])], true);
    } catch (e2) {}
}

function _applySlideMotion(trackItem, categoryName, seq, slideAnchor) {
    if (!trackItem || categoryName === 'Be Vegan Keep Peace') return;

    var motion = _findComponentByMatchName(trackItem, 'ADBE Motion', 'Motion');
    if (!motion) return;

    var position = _findComponentProperty(motion, 'ADBE Position', 'Position');
    var scale = _findComponentProperty(motion, 'ADBE Scale', 'Scale');
    var preset = _getMotionPreset(slideAnchor);

    _setMotionPosition(position, preset.position, seq);
    _setComponentPropertyValue(scale, preset.scale);
}

function _buildPlacementContext(seq, batches, targetTrackNumber, ignoreV1) {
    var allPlacements = [];
    var usedTimelineLengthSeconds = _getSequenceMaxEndSeconds(seq);

    if (usedTimelineLengthSeconds <= 0) {
        usedTimelineLengthSeconds = seq.end ? seq.end.seconds : 0;
    }
    if (usedTimelineLengthSeconds <= 0) {
        usedTimelineLengthSeconds = 1;
    }

    for (var i = 0; i < batches.length; i++) {
        var batch = batches[i];
        var files = batch.files || [];
        var languageDetails = batch.languageDetails || [];
        for (var f = 0; f < files.length; f++) {
            allPlacements.push({
                placementIndex: allPlacements.length,
                batchIndex: i,
                fileIndex: f,
                categoryName: batch.categoryName,
                mediaPath: files[f],
                language: languageDetails[f] ? languageDetails[f].name : '',
                isEnglish: languageDetails[f] ? !!languageDetails[f].isEnglish : (f === 0),
                warning: batch.warning || ''
            });
        }
    }

    var placementWindow = _getPlacementWindow(seq, usedTimelineLengthSeconds);
    var blockedV1Ranges = ignoreV1 ? _clipRangesToWindow(_getTrackOccupiedRanges(seq.videoTracks[0]), placementWindow.start, placementWindow.end) : [];
    var availableRanges = ignoreV1 ? _getAvailableRanges(placementWindow.end, blockedV1Ranges) : [{ start: placementWindow.start, end: placementWindow.end }];
    if (ignoreV1) {
        availableRanges = _clipRangesToWindow(availableRanges, placementWindow.start, placementWindow.end);
    }

    var usableTimelineLengthSeconds = _getTotalRangeLength(availableRanges);
    if (usableTimelineLengthSeconds <= 0) {
        usableTimelineLengthSeconds = placementWindow.end - placementWindow.start;
        availableRanges = [{ start: placementWindow.start, end: placementWindow.end }];
    }

    var intervalSeconds = allPlacements.length > 0 ? usableTimelineLengthSeconds / allPlacements.length : 0;

    for (var p = 0; p < allPlacements.length; p++) {
        var desiredCompressedSeconds = (p + 1) * intervalSeconds;
        allPlacements[p].startSeconds = ignoreV1 ? _mapCompressedTimeToSequenceTime(desiredCompressedSeconds, availableRanges) : (placementWindow.start + desiredCompressedSeconds);
    }

    return {
        allPlacements: allPlacements,
        resolvedTrackNumber: targetTrackNumber,
        usedTimelineLengthSeconds: usedTimelineLengthSeconds,
        usableTimelineLengthSeconds: usableTimelineLengthSeconds,
        placementWindowStartSeconds: placementWindow.start,
        placementWindowEndSeconds: placementWindow.end,
        intervalSeconds: intervalSeconds,
        blockedV1Ranges: blockedV1Ranges
    };
}

function newPeaceMakerPreviewPlacementFrames(payloadJson) {
    try {
        var payload = _npmParse(payloadJson);
        var batches = payload.batches || [];
        var targetTrackNumber = parseInt(payload.targetTrack, 10);
        var ignoreV1 = !!payload.ignoreV1;
        var analysisDir = payload.analysisDir || '';
        var exportFrames = !!payload.exportFrames;
        var placementPlan = payload.placementPlan || null;

        if (!app.project) {
            return _npmJson({ ok: false, error: 'No open project.' });
        }
        if (!app.project.activeSequence) {
            return _npmJson({ ok: false, error: 'No active sequence.' });
        }
        if (!batches.length) {
            return _npmJson({ ok: false, error: 'No category batches were provided.' });
        }
        if (!targetTrackNumber || targetTrackNumber < 1) {
            return _npmJson({ ok: false, error: 'Invalid target track.' });
        }
        if (exportFrames && (!analysisDir || !_ensureFolder(analysisDir))) {
            return _npmJson({ ok: false, error: 'Could not create the frame analysis folder.' });
        }

        var seq = app.project.activeSequence;
        var rootItem = app.project.rootItem;
        var slidesBin = _findOrCreateBin(rootItem, 'Slides');
        var context = _buildPlacementContext(seq, batches, targetTrackNumber, ignoreV1);
        if (placementPlan && placementPlan.placements && placementPlan.placements.length === context.allPlacements.length) {
            for (var pp = 0; pp < context.allPlacements.length; pp++) {
                if (typeof placementPlan.placements[pp].startSeconds === 'number') {
                    context.allPlacements[pp].startSeconds = placementPlan.placements[pp].startSeconds;
                }
            }
        }
        var placements = [];
        var exportFailures = [];

        for (var b = 0; b < batches.length; b++) {
            var previewBatch = batches[b];
            var previewSubBin = _findOrCreateBin(slidesBin, previewBatch.categoryName);
            for (var x = 0; x < context.allPlacements.length; x++) {
                if (context.allPlacements[x].batchIndex === b) {
                    context.allPlacements[x].subBin = previewSubBin;
                }
            }
        }

        for (var i = 0; i < context.allPlacements.length; i++) {
            var placement = context.allPlacements[i];
            var projectItem = _ensureProjectItemInBin(placement.subBin, placement.mediaPath);
            var clipDurationSeconds = _getProjectItemDurationSeconds(projectItem);
            if (clipDurationSeconds <= 0) {
                clipDurationSeconds = 1;
            }
            placement.clipDurationSeconds = clipDurationSeconds;
            var framePaths = [];
            var sampleTimes = [];
            var exportErrors = [];
            if (exportFrames) {
                var sampleExport = _exportPlacementSampleFrames(seq, placement.startSeconds, clipDurationSeconds, analysisDir, _sanitizeFileStem((i + 1) + '_' + placement.categoryName + '_' + placement.language));
                framePaths = sampleExport.framePaths;
                sampleTimes = sampleExport.sampleTimes;
                exportErrors = sampleExport.exportErrors;
                if (exportErrors.length) {
                    exportFailures.push('[' + (i + 1) + '] ' + exportErrors.join(' | '));
                }
            }
            var anyFrameExported = false;
            for (var fp = 0; fp < framePaths.length; fp++) {
                if (framePaths[fp]) {
                    anyFrameExported = true;
                    break;
                }
            }

            placements.push({
                placementIndex: placement.placementIndex,
                batchIndex: placement.batchIndex,
                fileIndex: placement.fileIndex,
                categoryName: placement.categoryName,
                mediaPath: placement.mediaPath,
                language: placement.language,
                isEnglish: placement.isEnglish,
                startSeconds: placement.startSeconds,
                clipDurationSeconds: clipDurationSeconds,
                sampleTimes: sampleTimes,
                framePaths: framePaths,
                framePath: framePaths.length ? framePaths[0] : '',
                frameExported: anyFrameExported
            });
        }

        return _npmJson({
            ok: true,
            placementPlan: {
                placements: placements,
                targetTrack: context.resolvedTrackNumber,
                usedTimelineLengthSeconds: context.usedTimelineLengthSeconds,
                usableTimelineLengthSeconds: context.usableTimelineLengthSeconds,
                placementWindowStartSeconds: context.placementWindowStartSeconds,
                placementWindowEndSeconds: context.placementWindowEndSeconds,
                intervalSeconds: context.intervalSeconds,
                blockedV1Ranges: context.blockedV1Ranges || []
            },
            frameExportFailures: exportFailures,
            note: exportFrames ? 'Preview frames were exported from the visible active sequence frame at each planned slide time.' : 'Placement times were calculated without exporting preview frames.'
        });
    } catch (err) {
        return _npmJson({ ok: false, error: err.toString() });
    }
}

function newPeaceMakerPreviewSinglePlacement(payloadJson) {
    try {
        var payload = _npmParse(payloadJson);
        var analysisDir = payload.analysisDir || '';
        var startSeconds = typeof payload.startSeconds === 'number' ? payload.startSeconds : 0;
        var clipDurationSeconds = typeof payload.clipDurationSeconds === 'number' ? payload.clipDurationSeconds : 1;
        var categoryName = payload.categoryName || '';
        var language = payload.language || '';
        var placementIndex = typeof payload.placementIndex === 'number' ? payload.placementIndex : 0;

        if (!app.project || !app.project.activeSequence) {
            return _npmJson({ ok: false, error: 'No active sequence.' });
        }
        if (!analysisDir || !_ensureFolder(analysisDir)) {
            return _npmJson({ ok: false, error: 'Could not create the frame analysis folder.' });
        }

        var seq = app.project.activeSequence;
        var sampleExport = _exportPlacementSampleFrames(seq, startSeconds, clipDurationSeconds, analysisDir, _sanitizeFileStem((placementIndex + 1) + '_' + categoryName + '_' + language));
        var anyFrameExported = false;
        for (var i = 0; i < sampleExport.framePaths.length; i++) {
            if (sampleExport.framePaths[i]) {
                anyFrameExported = true;
                break;
            }
        }

        return _npmJson({
            ok: true,
            placementPreview: {
                placementIndex: placementIndex,
                startSeconds: startSeconds,
                clipDurationSeconds: clipDurationSeconds,
                sampleTimes: sampleExport.sampleTimes,
                framePaths: sampleExport.framePaths,
                framePath: sampleExport.framePaths.length ? sampleExport.framePaths[0] : '',
                frameExported: anyFrameExported
            },
            exportErrors: sampleExport.exportErrors
        });
    } catch (err) {
        return _npmJson({ ok: false, error: err.toString() });
    }
}

function newPeaceMakerImportAndPlaceMulti(payloadJson) {
    try {
        var payload = _npmParse(payloadJson);
        var batches = payload.batches || [];
        var targetTrackNumber = parseInt(payload.targetTrack, 10);
        var ignoreV1 = !!payload.ignoreV1;
        var slideAnchor = payload.slideAnchor || 'top-right';
        var placementPlan = payload.placementPlan || null;

        if (!app.project) {
            return _npmJson({ ok: false, error: 'No open project.' });
        }
        if (!app.project.activeSequence) {
            return _npmJson({ ok: false, error: 'No active sequence.' });
        }
        if (!batches.length) {
            return _npmJson({ ok: false, error: 'No category batches were provided.' });
        }
        if (!targetTrackNumber || targetTrackNumber < 1) {
            return _npmJson({ ok: false, error: 'Invalid target track.' });
        }

        var rootItem = app.project.rootItem;
        var slidesBin = _findOrCreateBin(rootItem, 'Slides');
        var seq = app.project.activeSequence;
        var context = _buildPlacementContext(seq, batches, targetTrackNumber, ignoreV1);
        var allPlacements = context.allPlacements;
        var results = [];

        for (var i = 0; i < batches.length; i++) {
            var batch = batches[i];
            var subBin = _findOrCreateBin(slidesBin, batch.categoryName);
            var files = batch.files || [];
            app.project.importFiles(files, false, subBin, false);
            for (var p = 0; p < allPlacements.length; p++) {
                if (allPlacements[p].batchIndex === i) {
                    allPlacements[p].subBin = subBin;
                }
            }
        }

        var totalCount = allPlacements.length;
        var resolvedTrackNumber = context.resolvedTrackNumber;

        var trackIndex = resolvedTrackNumber - 1;
        while (seq.videoTracks.numTracks <= trackIndex) {
            seq.videoTracks.addTrack();
        }
        var destTrack = seq.videoTracks[trackIndex];

        for (var j = 0; j < allPlacements.length; j++) {
            var placement = allPlacements[j];
            var projectItem = _findProjectItemByMediaPath(placement.subBin, placement.mediaPath);
            if (!projectItem) {
                continue;
            }
            _setProjectItemLabel(projectItem, placement.isEnglish);
            var when = new Time();
            if (placementPlan && placementPlan.placements && placementPlan.placements.length > j && typeof placementPlan.placements[j].startSeconds === 'number') {
                when.seconds = placementPlan.placements[j].startSeconds;
                placement.resolvedAnchor = placementPlan.placements[j].resolvedAnchor || slideAnchor;
            } else if (typeof placement.startSeconds === 'number') {
                when.seconds = placement.startSeconds;
            }
            destTrack.overwriteClip(projectItem, when);
            var insertedClip = _findTrackItemByStartSeconds(destTrack, when.seconds);
            _applySlideMotion(insertedClip, placement.categoryName, seq, placement.resolvedAnchor || slideAnchor);
        }

        for (var k = 0; k < batches.length; k++) {
            var batchResult = batches[k];
            results.push({
                categoryName: batchResult.categoryName,
                targetTrack: resolvedTrackNumber,
                importedCount: (batchResult.files || []).length,
                placedCount: (batchResult.files || []).length,
                usedTimelineLengthSeconds: context.usedTimelineLengthSeconds,
                usableTimelineLengthSeconds: context.usableTimelineLengthSeconds,
                placementWindowStartSeconds: context.placementWindowStartSeconds,
                placementWindowEndSeconds: context.placementWindowEndSeconds,
                intervalSeconds: context.intervalSeconds,
                warning: batchResult.warning || ''
            });
        }

        return _npmJson({
            ok: true,
            results: results,
            targetTrack: resolvedTrackNumber,
            totalPlacedCount: totalCount,
            usedTimelineLengthSeconds: context.usedTimelineLengthSeconds,
            usableTimelineLengthSeconds: context.usableTimelineLengthSeconds,
            placementWindowStartSeconds: context.placementWindowStartSeconds,
            placementWindowEndSeconds: context.placementWindowEndSeconds,
            intervalSeconds: context.intervalSeconds,
            note: placementPlan ? 'All selected clips were placed using panel-resolved safe times on a single video track.' : 'All selected clips were placed on a single video track using one shared interval calculated as used timeline length divided by total selected slides.'
        });
    } catch (err) {
        return _npmJson({ ok: false, error: err.toString() });
    }
}
