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

var _NPM_LABEL_ENGLISH = 4;
var _NPM_LABEL_NON_ENGLISH = 8;
var _NPM_MOTION_POSITION = [1795.0, 336.0];
var _NPM_MOTION_SCALE = 66.0;

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

function _setMotionPositionAtClipStart(property, trackItem, xy, seq) {
    if (!property || !trackItem || !trackItem.start || !xy || xy.length < 2) return;
    if (typeof property.setTimeVarying !== 'function' || typeof property.addKey !== 'function' || typeof property.setValueAtKey !== 'function') return;
    var relativeXY = _getSequenceRelativePosition(seq, xy);

    var keyTime = new Time();
    keyTime.seconds = trackItem.start.seconds;

    try {
        property.setTimeVarying(true);
    } catch (e) {}

    try {
        var keys = property.getKeys ? property.getKeys() : null;
        if (keys && keys.length) {
            for (var i = 0; i < keys.length; i++) {
                try {
                    if (property.removeKey) {
                        property.removeKey(keys[i]);
                    }
                } catch (removeErr) {}
            }
        }
    } catch (e2) {}

    try {
        property.addKey(keyTime);
    } catch (e3) {}

    try {
        property.setValueAtKey(keyTime, [relativeXY[0], relativeXY[1]], 1);
    } catch (e4) {
        try {
            property.setValueAtKey(keyTime, [parseFloat(relativeXY[0]), parseFloat(relativeXY[1])], 1);
        } catch (e5) {}
    }
}

function _applySlideMotion(trackItem, categoryName, seq) {
    if (!trackItem || categoryName === 'Be Vegan Keep Peace') return;

    var motion = _findComponentByMatchName(trackItem, 'ADBE Motion', 'Motion');
    if (!motion) return;

    var position = _findComponentProperty(motion, 'ADBE Position', 'Position');
    var scale = _findComponentProperty(motion, 'ADBE Scale', 'Scale');

    _setMotionPosition(position, _NPM_MOTION_POSITION, seq);
    _setMotionPositionAtClipStart(position, trackItem, _NPM_MOTION_POSITION, seq);
    _setComponentPropertyValue(scale, _NPM_MOTION_SCALE);
}

function newPeaceMakerImportAndPlaceMulti(payloadJson) {
    try {
        var payload = _npmParse(payloadJson);
        var batches = payload.batches || [];
        var targetTrackNumber = parseInt(payload.targetTrack, 10);
        var ignoreV1 = !!payload.ignoreV1;

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
        var allPlacements = [];
        var results = [];

        for (var i = 0; i < batches.length; i++) {
            var batch = batches[i];
            var subBin = _findOrCreateBin(slidesBin, batch.categoryName);
            var files = batch.files || [];
            var languageDetails = batch.languageDetails || [];
            app.project.importFiles(files, false, subBin, false);
            for (var f = 0; f < files.length; f++) {
                allPlacements.push({
                    categoryName: batch.categoryName,
                    mediaPath: files[f],
                    language: languageDetails[f] ? languageDetails[f].name : '',
                    isEnglish: languageDetails[f] ? !!languageDetails[f].isEnglish : (f === 0),
                    subBin: subBin,
                    warning: batch.warning || ''
                });
            }
        }

        var usedTimelineLengthSeconds = _getSequenceMaxEndSeconds(seq);
        if (usedTimelineLengthSeconds <= 0) {
            usedTimelineLengthSeconds = seq.end ? seq.end.seconds : 0;
        }
        if (usedTimelineLengthSeconds <= 0) {
            usedTimelineLengthSeconds = 1;
        }

        var placementWindow = _getPlacementWindow(seq, usedTimelineLengthSeconds);
        var totalCount = allPlacements.length;
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
        var intervalSeconds = totalCount > 0 ? usableTimelineLengthSeconds / totalCount : 0;

        var resolvedTrackNumber = targetTrackNumber;

        var trackIndex = resolvedTrackNumber - 1;
        while (seq.videoTracks.numTracks <= trackIndex) {
            seq.videoTracks.addTrack();
        }
        var destTrack = seq.videoTracks[trackIndex];

        for (var p = 0; p < allPlacements.length; p++) {
            var placement = allPlacements[p];
            var projectItem = _findProjectItemByMediaPath(placement.subBin, placement.mediaPath);
            if (!projectItem) {
                continue;
            }
            _setProjectItemLabel(projectItem, placement.isEnglish);
            var when = new Time();
            var desiredCompressedSeconds = (p + 1) * intervalSeconds;
            when.seconds = ignoreV1 ? _mapCompressedTimeToSequenceTime(desiredCompressedSeconds, availableRanges) : desiredCompressedSeconds;
            if (!ignoreV1) {
                when.seconds = placementWindow.start + desiredCompressedSeconds;
            }
            destTrack.overwriteClip(projectItem, when);
            var insertedClip = _findTrackItemByStartSeconds(destTrack, when.seconds);
            _applySlideMotion(insertedClip, placement.categoryName, seq);
        }

        for (var j = 0; j < batches.length; j++) {
            var batchResult = batches[j];
            results.push({
                categoryName: batchResult.categoryName,
                targetTrack: resolvedTrackNumber,
                importedCount: (batchResult.files || []).length,
                placedCount: (batchResult.files || []).length,
                usedTimelineLengthSeconds: usedTimelineLengthSeconds,
                usableTimelineLengthSeconds: usableTimelineLengthSeconds,
                placementWindowStartSeconds: placementWindow.start,
                placementWindowEndSeconds: placementWindow.end,
                intervalSeconds: intervalSeconds,
                warning: batchResult.warning || ''
            });
        }

        return _npmJson({
            ok: true,
            results: results,
            targetTrack: resolvedTrackNumber,
            totalPlacedCount: totalCount,
            usedTimelineLengthSeconds: usedTimelineLengthSeconds,
            usableTimelineLengthSeconds: usableTimelineLengthSeconds,
            placementWindowStartSeconds: placementWindow.start,
            placementWindowEndSeconds: placementWindow.end,
            intervalSeconds: intervalSeconds,
            note: 'All selected clips were placed on a single video track using one shared interval calculated as used timeline length divided by total selected slides.'
        });
    } catch (err) {
        return _npmJson({ ok: false, error: err.toString() });
    }
}
