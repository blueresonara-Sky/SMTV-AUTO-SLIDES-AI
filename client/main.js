(function () {
  'use strict';

  var cs = new CSInterface();
  var fs = require('fs');
  var path = require('path');
  var os = require('os');
  var https = require('https');
  var childProcess = require('child_process');

  var folderPicker = document.getElementById('folderPicker');
  var browseBtn = document.getElementById('browseBtn');
  var folderPathInput = document.getElementById('folderPath');
  var runBtn = document.getElementById('runBtn');
  var slideCountInput = document.getElementById('slideCount');
  var targetTrackInput = document.getElementById('targetTrack');
  var ignoreV1Input = document.getElementById('ignoreV1');
  var slideAnchorInput = document.getElementById('slideAnchor');
  var avoidFacesInput = document.getElementById('avoidFaces');
  var installedVersionEl = document.getElementById('installedVersion');
  var latestVersionEl = document.getElementById('latestVersion');
  var installUpdateBtn = document.getElementById('installUpdateBtn');
  var updateStatusEl = document.getElementById('updateStatus');
  var statusEl = document.getElementById('status');
  var chosenTitleEl = document.getElementById('chosenTitle');
  var chosenLanguagesEl = document.getElementById('chosenLanguages');
  var updateModalEl = document.getElementById('updateModal');
  var updateModalTitleEl = document.getElementById('updateModalTitle');
  var updateModalBodyEl = document.getElementById('updateModalBody');
  var updateModalCancelBtn = document.getElementById('updateModalCancelBtn');
  var updateModalOkBtn = document.getElementById('updateModalOkBtn');

  var selectedRootFolder = '';
  var extensionRoot = '';
  var manifestPath = '';
  var trackingDir = path.join(os.homedir(), '.new-peace-maker');
  var trackingFile = path.join(trackingDir, 'usage-history.json');
  var updateInstallStatusFile = path.join(trackingDir, 'update-install-status.json');
  var TEST_UPDATE_FLAG_FILE = 'smtv-auto-slides-test-updates.flag';
  var UPDATE_CHANNEL_STORAGE_KEY = 'smtvAutoSlides_updateChannel';
  var categoryOrder = ['NEW PEACE MAKER', 'Be Vegan Keep Peace', 'Forgiveness', 'Save the Earth', 'Veganism'];
  var ignoredFolderNames = { 'AFTERCODECS HAP ALPHA': true };
  var updateRepo = 'blueresonara-Sky/SMTV-AUTO-SLIDES-AI';
  var updateState = {
    installedVersion: '',
    latestVersion: '',
    latestRelease: null,
    checking: false,
    installing: false
  };

  // Neural face-detection state (face-api.js / TinyFaceDetector)
  var faceApiReady = false;
  var faceApiInitError = null;

  // OCR text-detection state (OCRAD.js)
  var ocradReady = false;

  function defaultTracking() {
    return {
      categories: {},
      usedLanguagesGlobalCycle: [],
      ignoredFolders: ['AFTERCODECS HAP ALPHA'],
      settings: {
        rootFolder: '',
        slideCount: 6,
        targetTrack: 9,
        ignoreV1: false,
        slideAnchor: 'top-right',
        avoidFaces: true,
        lastUpdateCheckAt: '',
        lastAvailableVersion: '',
        pendingUpdateVersion: '',
        pendingUpdateName: '',
        pendingUpdateNotes: '',
      }
    };
  }

  function log(msg) {
    statusEl.textContent += '\n' + msg;
    statusEl.scrollTop = statusEl.scrollHeight;
  }

  // Format seconds as m:ss:ff at 29.97 fps  e.g. 203.7 → "3:23:21"
  function secToMS(s) {
    var fps    = 29.97;
    var neg    = s < 0;
    var abs    = Math.abs(s);
    var totalF = Math.round(abs * fps);
    var ff     = totalF % Math.round(fps);          // frames (0-29)
    var totS   = Math.floor(totalF / Math.round(fps));
    var ss     = totS % 60;
    var m      = Math.floor(totS / 60);
    return (neg ? '-' : '') + m + ':' + (ss < 10 ? '0' : '') + ss + ':' + (ff < 10 ? '0' : '') + ff;
  }

  function setStatus(msg) {
    statusEl.textContent = msg;
  }

  function shuffle(arr) {
    var a = arr.slice();
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var temp = a[i];
      a[i] = a[j];
      a[j] = temp;
    }
    return a;
  }

  function ensureTrackingFile() {
    if (!fs.existsSync(trackingDir)) {
      fs.mkdirSync(trackingDir, { recursive: true });
    }
    if (!fs.existsSync(trackingFile)) {
      fs.writeFileSync(trackingFile, JSON.stringify(defaultTracking(), null, 2), 'utf8');
    }
  }

  function resolveExtensionRoot() {
    try {
      if (cs && typeof cs.getSystemPath === 'function' && typeof SystemPath !== 'undefined' && typeof SystemPath.EXTENSION !== 'undefined') {
        var cepExtensionPath = cs.getSystemPath(SystemPath.EXTENSION);
        if (cepExtensionPath && fs.existsSync(cepExtensionPath)) {
          return cepExtensionPath;
        }
      }
    } catch (e) {}

    try {
      if (typeof window !== 'undefined' && window.location && window.location.pathname) {
        var pathname = decodeURIComponent(window.location.pathname).replace(/^\/([A-Za-z]:\/)/, '$1');
        var htmlPath = pathname.replace(/\//g, path.sep);
        var fromLocation = path.resolve(path.dirname(htmlPath), '..');
        if (fromLocation && fs.existsSync(fromLocation)) {
          return fromLocation;
        }
      }
    } catch (e1) {}

    try {
      var fallbackPath = path.resolve(__dirname, '..');
      if (fallbackPath && fs.existsSync(fallbackPath)) {
        return fallbackPath;
      }
    } catch (e2) {}

    return '';
  }

  function loadTracking() {
    ensureTrackingFile();
    try {
      var parsed = JSON.parse(fs.readFileSync(trackingFile, 'utf8'));
      var base = defaultTracking();
      parsed.categories = parsed.categories || {};
      parsed.usedLanguagesGlobalCycle = Array.isArray(parsed.usedLanguagesGlobalCycle) ? parsed.usedLanguagesGlobalCycle : [];
      parsed.usedLanguagesGlobalCycle = parsed.usedLanguagesGlobalCycle
        .map(function (lang) { return canonicalizeLanguageName(lang); })
        .filter(function (lang, index, arr) { return !!lang && arr.indexOf(lang) === index; });
      parsed.ignoredFolders = Array.isArray(parsed.ignoredFolders) ? parsed.ignoredFolders : base.ignoredFolders;
      parsed.settings = parsed.settings || base.settings;
      if (typeof parsed.settings.slideCount === 'undefined') parsed.settings.slideCount = base.settings.slideCount;
      if (typeof parsed.settings.targetTrack === 'undefined') parsed.settings.targetTrack = base.settings.targetTrack;
      if (typeof parsed.settings.rootFolder === 'undefined') parsed.settings.rootFolder = base.settings.rootFolder;
      if (typeof parsed.settings.ignoreV1 === 'undefined') parsed.settings.ignoreV1 = base.settings.ignoreV1;
      if (typeof parsed.settings.slideAnchor === 'undefined') parsed.settings.slideAnchor = base.settings.slideAnchor;
        if (typeof parsed.settings.avoidFaces === 'undefined') parsed.settings.avoidFaces = base.settings.avoidFaces;
        if (typeof parsed.settings.lastUpdateCheckAt === 'undefined') parsed.settings.lastUpdateCheckAt = base.settings.lastUpdateCheckAt;
        if (typeof parsed.settings.lastAvailableVersion === 'undefined') parsed.settings.lastAvailableVersion = base.settings.lastAvailableVersion;
        if (typeof parsed.settings.pendingUpdateVersion === 'undefined') parsed.settings.pendingUpdateVersion = base.settings.pendingUpdateVersion;
        if (typeof parsed.settings.pendingUpdateName === 'undefined') parsed.settings.pendingUpdateName = base.settings.pendingUpdateName;
        if (typeof parsed.settings.pendingUpdateNotes === 'undefined') parsed.settings.pendingUpdateNotes = base.settings.pendingUpdateNotes;
        return parsed;
    } catch (e) {
      return defaultTracking();
    }
  }

  function saveTracking(data) {
    ensureTrackingFile();
    fs.writeFileSync(trackingFile, JSON.stringify(data, null, 2), 'utf8');
  }

  function loadUpdateInstallStatus() {
    ensureTrackingFile();
    try {
      if (!fs.existsSync(updateInstallStatusFile)) return null;
      return JSON.parse(fs.readFileSync(updateInstallStatusFile, 'utf8'));
    } catch (e) {
      return null;
    }
  }

  function saveUpdateInstallStatus(data) {
    ensureTrackingFile();
    fs.writeFileSync(updateInstallStatusFile, JSON.stringify(data || {}, null, 2), 'utf8');
  }

  function clearUpdateInstallStatus() {
    ensureTrackingFile();
    try {
      if (fs.existsSync(updateInstallStatusFile)) {
        fs.unlinkSync(updateInstallStatusFile);
      }
    } catch (e) {}
  }

  function savePendingUpdateInfo(version, name, notes) {
    var tracking = loadTracking();
    tracking.settings.pendingUpdateVersion = version || '';
    tracking.settings.pendingUpdateName = name || '';
    tracking.settings.pendingUpdateNotes = notes || '';
    saveTracking(tracking);
  }

  function clearPendingUpdateInfo() {
    savePendingUpdateInfo('', '', '');
  }

  function getPendingUpdateInfo() {
    var tracking = loadTracking();
    return {
      version: tracking.settings.pendingUpdateVersion || '',
      name: tracking.settings.pendingUpdateName || '',
      notes: tracking.settings.pendingUpdateNotes || ''
    };
  }

  function getReleaseNotes(release) {
    var body = release && release.body ? String(release.body) : '';
    var name = release && (release.name || release.tag_name) ? String(release.name || release.tag_name) : '';
    var notes = body.replace(/\r/g, '').trim();
    if (!notes) {
      notes = name ? ('Release: ' + name) : 'No release notes were provided for this update.';
    }
    if (notes.length > 4000) {
      notes = notes.substring(0, 4000).replace(/\s+\S*$/, '') + '\n\n...';
    }
    return notes;
  }

  function getPopupReleaseNotes(release) {
    var notes = getReleaseNotes(release);
    var parts = notes.split(/\n---\n/);
    var popupNotes = parts[0] ? parts[0].trim() : notes;
    return popupNotes || notes;
  }

  function setModalOpen(isOpen) {
    if (!updateModalEl) return;
    updateModalEl.className = isOpen ? 'modal-backdrop is-open' : 'modal-backdrop';
  }

  function showUpdateModal(title, message, options) {
    return new Promise(function (resolve) {
      if (!updateModalEl || !updateModalTitleEl || !updateModalBodyEl || !updateModalOkBtn || !updateModalCancelBtn) {
        if (options && options.confirm) {
          resolve(window.confirm(title + '\n\n' + message));
        } else {
          window.alert(title + '\n\n' + message);
          resolve(true);
        }
        return;
      }

      updateModalTitleEl.textContent = title;
      updateModalBodyEl.textContent = message;
      updateModalOkBtn.textContent = options && options.okText ? options.okText : 'OK';
      updateModalCancelBtn.textContent = options && options.cancelText ? options.cancelText : 'Cancel';
      updateModalCancelBtn.style.display = options && options.confirm ? 'inline-block' : 'none';

      function cleanup(result) {
        updateModalOkBtn.removeEventListener('click', onOk);
        updateModalCancelBtn.removeEventListener('click', onCancel);
        updateModalEl.removeEventListener('click', onBackdrop);
        setModalOpen(false);
        resolve(result);
      }

      function onOk() { cleanup(true); }
      function onCancel() { cleanup(false); }
      function onBackdrop(evt) {
        if (evt.target === updateModalEl && options && options.confirm) {
          cleanup(false);
        }
      }

      updateModalOkBtn.addEventListener('click', onOk);
      updateModalCancelBtn.addEventListener('click', onCancel);
      updateModalEl.addEventListener('click', onBackdrop);
      setModalOpen(true);
    });
  }

  function buildUpdateNotesMessage(release, prefix, options) {
    var message = prefix ? String(prefix).replace(/\s+$/, '') : '';
    var notes = options && options.popupSummary ? getPopupReleaseNotes(release) : getReleaseNotes(release);
    if (message) {
      message += '\n\n';
    }
    message += notes;
    return message;
  }

  function persistSettings() {
    var tracking = loadTracking();
    tracking.settings.rootFolder = selectedRootFolder || '';
    tracking.settings.slideCount = parseInt(slideCountInput.value, 10) || 6;
    tracking.settings.targetTrack = parseInt(targetTrackInput.value, 10) || 9;
    tracking.settings.ignoreV1 = !!(ignoreV1Input && ignoreV1Input.checked);
    tracking.settings.slideAnchor = slideAnchorInput ? String(slideAnchorInput.value || 'top-right') : 'top-right';
    tracking.settings.avoidFaces = !!(avoidFacesInput && avoidFacesInput.checked);
    saveTracking(tracking);
  }

  function restoreSettings() {
    var tracking = loadTracking();
    if (tracking.settings.rootFolder) {
      selectedRootFolder = tracking.settings.rootFolder;
      folderPathInput.value = selectedRootFolder;
    }
    slideCountInput.value = tracking.settings.slideCount || 6;
    targetTrackInput.value = tracking.settings.targetTrack || 9;
    ignoreV1Input.checked = !!tracking.settings.ignoreV1;
    slideAnchorInput.value = tracking.settings.slideAnchor || 'top-right';
    avoidFacesInput.checked = tracking.settings.avoidFaces !== false;
  }

  function setUpdateStatus(msg) {
    updateStatusEl.textContent = msg;
  }

  function isLocalPrereleaseReinstallAvailable(release) {
    return !!(
      isTestUpdateChannelEnabled() &&
      release &&
      release.prerelease &&
      getReleaseZipAsset(release) &&
      compareVersions(getReleaseVersion(release), updateState.installedVersion) === 0
    );
  }

  function setUpdateUiState() {
    var hasUpdate = !!updateState.latestRelease && (
      compareVersions(updateState.latestVersion, updateState.installedVersion) > 0 ||
      isLocalPrereleaseReinstallAvailable(updateState.latestRelease)
    );
    installedVersionEl.textContent = updateState.installedVersion || '-';
    latestVersionEl.textContent = updateState.latestVersion || '-';
    installUpdateBtn.disabled = updateState.checking || updateState.installing || !updateState.latestRelease;
    installUpdateBtn.hidden = !hasUpdate;
    if (installUpdateBtn.classList) {
      installUpdateBtn.classList.toggle('update-available', hasUpdate);
    }
  }

  function readManifestVersion(filePath) {
    try {
      var manifestXml = fs.readFileSync(filePath, 'utf8');
      var match = manifestXml.match(/ExtensionBundleVersion="([^"]+)"/);
      return match ? match[1] : '';
    } catch (e) {
      return '';
    }
  }

  function readManifestBundleId(filePath) {
    try {
      var manifestXml = fs.readFileSync(filePath, 'utf8');
      var match = manifestXml.match(/ExtensionBundleId="([^"]+)"/);
      return match ? match[1] : '';
    } catch (e) {
      return '';
    }
  }

  function normalizeVersion(version) {
    return String(version || '').trim().replace(/^v/i, '');
  }

  function compareVersions(a, b) {
    var aParts = normalizeVersion(a).split('.');
    var bParts = normalizeVersion(b).split('.');
    var maxLen = Math.max(aParts.length, bParts.length);
    for (var i = 0; i < maxLen; i++) {
      var aNum = parseInt(aParts[i] || '0', 10);
      var bNum = parseInt(bParts[i] || '0', 10);
      if (isNaN(aNum)) aNum = 0;
      if (isNaN(bNum)) bNum = 0;
      if (aNum > bNum) return 1;
      if (aNum < bNum) return -1;
    }
    return 0;
  }

  function persistUpdateInfo(latestVersion) {
    var tracking = loadTracking();
    tracking.settings.lastUpdateCheckAt = new Date().toISOString();
    tracking.settings.lastAvailableVersion = latestVersion || '';
    saveTracking(tracking);
  }

  function ensureDirExists(dirPath) {
    if (!dirPath) return;
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath, { recursive: true });
    }
  }

  function cleanupDir(dirPath) {
    if (!dirPath || !fs.existsSync(dirPath)) return;
    try {
      if (typeof fs.rmSync === 'function') {
        fs.rmSync(dirPath, { recursive: true, force: true });
        return;
      }
    } catch (e) {}

    try {
      fs.readdirSync(dirPath).forEach(function (entry) {
        var entryPath = path.join(dirPath, entry);
        var stat = fs.statSync(entryPath);
        if (stat.isDirectory()) {
          cleanupDir(entryPath);
        } else {
          fs.unlinkSync(entryPath);
        }
      });
      fs.rmdirSync(dirPath);
    } catch (e1) {}
  }

  function createFaceAnalysisTempDir() {
    var dirPath = path.join(os.tmpdir(), 'smtv-auto-slides-face-' + Date.now() + '-' + Math.floor(Math.random() * 100000));
    ensureDirExists(dirPath);
    return dirPath;
  }

  function loadImageFromFile(filePath, callback) {
    if (!filePath || !fs.existsSync(filePath)) {
      callback(new Error('Frame file was not found.'));
      return;
    }

    fs.readFile(filePath, function (err, buffer) {
      if (err) {
        callback(err);
        return;
      }

      var img = new Image();
      img.onload = function () {
        callback(null, img);
      };
      img.onerror = function () {
        callback(new Error('The exported frame could not be loaded.'));
      };
      img.src = 'data:image/png;base64,' + buffer.toString('base64');
    });
  }

  function getLuminance(r, g, b) {
    return (0.299 * r) + (0.587 * g) + (0.114 * b);
  }

  function isLikelySkinPixel(r, g, b) {
    var max = Math.max(r, g, b);
    var min = Math.min(r, g, b);
    var delta = max - min;
    var cb = 128 - (0.168736 * r) - (0.331264 * g) + (0.5 * b);
    var cr = 128 + (0.5 * r) - (0.418688 * g) - (0.081312 * b);
    // delta 20–110: two-sided saturation gate.
    //   delta < 20 → near-neutral grey/white (clouds, rain, sky, white walls) — NOT skin
    //   delta > 110 → hyper-saturated orange/amber (sunsets, golden light) — NOT skin
    //   Real skin across all ethnicities has delta ~20–100.
    var rgbRule = r > 45 && g > 20 && b > 10 && delta >= 20 && delta < 110 && r > b;
    // YCbCr range calibrated for diverse skin tones
    var yCbCrRule = cb >= 70 && cb <= 140 && cr >= 120 && cr <= 185;
    return rgbRule && yCbCrRule;
  }

  function scoreAnchorRegion(ctx, width, height, anchorKey) {
    var regions = {
      // Regions are calibrated to the ACTUAL slide graphic footprint in the frame,
      // derived from the motion presets in host/main.jsx:
      //   top-left  → position [936, 372] at 66% scale on 1920×1080
      //              → center ≈ (48.75%, 34.4%) → graphic spans ≈ x:35–63%, y:17–51%
      //   top-right → position [1795, 336] at 66% scale on 1920×1080
      //              → center ≈ (93.5%, 31.1%) → graphic spans ≈ x:79–100%, y:14–48%
      //
      // Each region matches the actual slide badge footprint measured from broadcast frames.
      //   top-left : globe+text badge  x=15–43%, y=4–13%
      //   top-right: text badge        x=66–100%, y=2–9%
      'top-left':  { x: 0.15, y: 0.04, w: 0.28, h: 0.09 },
      'top-right': { x: 0.66, y: 0.02, w: 0.34, h: 0.07 }
    };
    var region = regions[anchorKey] || regions['top-right'];
    var x = Math.max(0, Math.floor(width * region.x));
    var y = Math.max(0, Math.floor(height * region.y));
    var boxWidth = Math.max(1, Math.floor(width * region.w));
    var boxHeight = Math.max(1, Math.floor(height * region.h));
    if ((x + boxWidth) > width) boxWidth = width - x;
    if ((y + boxHeight) > height) boxHeight = height - y;

    var imageData = ctx.getImageData(x, y, boxWidth, boxHeight).data;
    var clearanceBandX = Math.max(0, Math.floor(width * (region.x + 0.015)));
    var clearanceBandY = Math.min(height - 1, y + boxHeight);
    var clearanceBandWidth = Math.max(1, Math.floor(width * Math.max(0.18, region.w - 0.03)));
    var clearanceBandHeight = Math.max(1, Math.floor(height * 0.16));
    if ((clearanceBandX + clearanceBandWidth) > width) clearanceBandWidth = width - clearanceBandX;
    if ((clearanceBandY + clearanceBandHeight) > height) clearanceBandHeight = height - clearanceBandY;
    var clearanceBandData = (clearanceBandWidth > 0 && clearanceBandHeight > 0)
      ? ctx.getImageData(clearanceBandX, clearanceBandY, clearanceBandWidth, clearanceBandHeight).data
      : null;
    var sampleWidth = boxWidth;
    var sampleHeight = boxHeight;
    var step = 2;
    var total = 0;
    var skin = 0;
    var topSkin = 0;
    var edges = 0;
    var texture = 0;
    var dark = 0;
    var lowerOccupied = 0;
    var lowerTotal = 0;
    var centralOccupied = 0;
    var centralTotal = 0;
    var clearanceOccupied = 0;
    var clearanceTotal = 0;
    var belowBandOccupied = 0;
    var belowBandTotal = 0;
    var lumSum = 0;
    var lumSqSum = 0;

    for (var yy = 0; yy < sampleHeight; yy += step) {
      for (var xx = 0; xx < sampleWidth; xx += step) {
        var idx = ((yy * sampleWidth) + xx) * 4;
        var r = imageData[idx];
        var g = imageData[idx + 1];
        var b = imageData[idx + 2];
        var lum = getLuminance(r, g, b);
        var saturation = Math.max(r, g, b) - Math.min(r, g, b);
        total++;
        lumSum += lum;
        lumSqSum += lum * lum;
        if (lum < 110) {
          dark++;
        }
        if (saturation > 26 && lum > 18 && lum < 245) {
          texture++;
        }
        var skinLike = isLikelySkinPixel(r, g, b);
        if (skinLike) {
          skin++;
          if (yy < (sampleHeight * 0.7)) {
            topSkin++;
          }
        }

        var occupiedPixel = skinLike || lum < 175 || saturation > 38;
        if (yy >= (sampleHeight * 0.35)) {
          lowerTotal++;
          if (occupiedPixel) {
            lowerOccupied++;
          }
        }
        if (xx >= (sampleWidth * 0.12) && xx <= (sampleWidth * 0.88)) {
          centralTotal++;
          if (occupiedPixel) {
            centralOccupied++;
          }
        }
        if (yy >= (sampleHeight * 0.72)) {
          clearanceTotal++;
          if (occupiedPixel) {
            clearanceOccupied++;
          }
        }

        if ((xx + step) < sampleWidth && (yy + step) < sampleHeight) {
          var rightIdx = ((yy * sampleWidth) + (xx + step)) * 4;
          var downIdx = ((((yy + step) * sampleWidth) + xx) * 4);
          var rightLum = getLuminance(imageData[rightIdx], imageData[rightIdx + 1], imageData[rightIdx + 2]);
          var downLum = getLuminance(imageData[downIdx], imageData[downIdx + 1], imageData[downIdx + 2]);
          var edgeMag = Math.abs(lum - rightLum) + Math.abs(lum - downLum);
          if (edgeMag > 55) {
            edges++;
          }
        }
      }
    }

    if (clearanceBandData) {
      for (var by = 0; by < clearanceBandHeight; by += step) {
        for (var bx = 0; bx < clearanceBandWidth; bx += step) {
          var bandIdx = ((by * clearanceBandWidth) + bx) * 4;
          var br = clearanceBandData[bandIdx];
          var bg = clearanceBandData[bandIdx + 1];
          var bb = clearanceBandData[bandIdx + 2];
          var bLum = getLuminance(br, bg, bb);
          var bSat = Math.max(br, bg, bb) - Math.min(br, bg, bb);
          var bandOccupied = isLikelySkinPixel(br, bg, bb) || bLum < 185 || bSat > 34;
          belowBandTotal++;
          if (bandOccupied) {
            belowBandOccupied++;
          }
        }
      }
    }

    var mean = total ? (lumSum / total) : 0;
    var variance = total ? Math.max(0, (lumSqSum / total) - (mean * mean)) : 0;
    var skinRatio = total ? (skin / total) : 0;
    var topSkinRatio = total ? (topSkin / total) : 0;
    var edgeRatio = total ? (edges / total) : 0;
    var textureRatio = total ? (texture / total) : 0;
    var darkRatio = total ? (dark / total) : 0;
    var lowerOccupiedRatio = lowerTotal ? (lowerOccupied / lowerTotal) : 0;
    var centralOccupiedRatio = centralTotal ? (centralOccupied / centralTotal) : 0;
    var clearanceRatio = clearanceTotal ? (clearanceOccupied / clearanceTotal) : 0;
    var belowBandRatio = belowBandTotal ? (belowBandOccupied / belowBandTotal) : 0;
    var varianceScore = Math.min(variance / 2500, 1);
    // Score is now PRIMARILY driven by skin detection.
    // Background busyness (texture, edges, dark pixels) gets much lower weight
    // so that a textured sofa or dark fabric does NOT outscore an actual face.
    var score = (skinRatio * 25) + (topSkinRatio * 20) + (clearanceRatio * 12) + (belowBandRatio * 14) + (lowerOccupiedRatio * 2) + (centralOccupiedRatio * 1.5) + (darkRatio * 0.5) + (edgeRatio * 0.5) + (textureRatio * 0.5) + (varianceScore * 0.3);

    return {
      anchor: anchorKey,
      score: score,
      skinRatio: skinRatio,
      topSkinRatio: topSkinRatio,
      darkRatio: darkRatio,
      lowerOccupiedRatio: lowerOccupiedRatio,
      centralOccupiedRatio: centralOccupiedRatio,
      clearanceRatio: clearanceRatio,
      belowBandRatio: belowBandRatio,
      edgeRatio: edgeRatio,
      textureRatio: textureRatio
    };
  }

  function detectTextInRegion(ctx, width, height, anchorKey) {
    // Detects broadcast TV text overlays (channel bugs, lower-thirds, badges, tickers).
    //
    // Strategy: TV text overlays always have a SOLID-COLOR BACKGROUND rectangle with
    // HIGH-CONTRAST text on top. Natural content (sky, stone, ruins) either lacks a
    // dominant background color OR lacks the high-contrast text pixels.
    //
    // For each horizontal strip of the region we:
    //   1. Build a luminance histogram and find the MODE (dominant color = background).
    //   2. Count "background pixels"  (lum within ±30 of mode).
    //   3. Count "text pixels"        (lum more than 70 away from mode).
    //   4. Flag strip as text-like if bgRatio > 60% AND textRatio > 5%.
    //
    // Returns fraction of strips that are text-like (0–1).
    //   Sky alone:          bgRatio ~95%, textRatio ~0%  → NOT text  ✓
    //   Stone/architecture: bgRatio ~25%, textRatio ~5%  → NOT text  ✓ (bgRatio too low)
    //   "Libya" badge:      bgRatio ~70%, textRatio ~20% → TEXT      ✓
    //   Ticker/super:       bgRatio ~65%, textRatio ~10% → TEXT      ✓
    var regionDefs = {
      'top-left':  { x: 0.15, y: 0.04, w: 0.28, h: 0.09 },
      'top-right': { x: 0.66, y: 0.02, w: 0.34, h: 0.07 }
    };
    var reg = regionDefs[anchorKey];
    if (!reg) return 0;

    var rx = Math.max(0, Math.floor(width  * reg.x));
    var ry = Math.max(0, Math.floor(height * reg.y));
    var rw = Math.min(Math.floor(width  * reg.w), width  - rx);
    var rh = Math.min(Math.floor(height * reg.h), height - ry);
    if (rw <= 0 || rh <= 0) return 0;

    var data   = ctx.getImageData(rx, ry, rw, rh).data;
    var step   = 2;    // sample every 2px for speed
    var stripH = 6;    // analyse 6-pixel-tall horizontal strips
    var textLikeStrips = 0;
    var totalStrips    = 0;

    for (var y0 = 0; y0 < rh; y0 += stripH) {
      var y1   = Math.min(y0 + stripH, rh);
      var lums = [];

      for (var yy = y0; yy < y1; yy += step) {
        for (var xx = 0; xx < rw; xx += step) {
          var idx = (yy * rw + xx) * 4;
          lums.push(0.299 * data[idx] + 0.587 * data[idx + 1] + 0.114 * data[idx + 2]);
        }
      }
      if (lums.length < 4) continue;

      // Find the mode luminance using 8-unit buckets
      var hist = [];
      var bi, bl;
      for (bi = 0; bi < 32; bi++) hist[bi] = 0;
      for (bl = 0; bl < lums.length; bl++) hist[Math.min(31, Math.floor(lums[bl] / 8))]++;
      var modeBucket = 0;
      for (bi = 1; bi < 32; bi++) { if (hist[bi] > hist[modeBucket]) modeBucket = bi; }
      var modeLum = modeBucket * 8 + 4;  // centre of dominant bucket

      // Score each pixel against the mode
      var bgPixels = 0, txtPixels = 0;
      for (bl = 0; bl < lums.length; bl++) {
        var d = Math.abs(lums[bl] - modeLum);
        if (d < 30) bgPixels++;   // close to dominant color = background
        if (d > 70) txtPixels++;  // far from dominant color = text on top
      }

      var bgRatio  = bgPixels  / lums.length;
      var txtRatio = txtPixels / lums.length;

      // Text strip: solid background (≥60%) AND visible high-contrast text (≥5%)
      if (bgRatio >= 0.60 && txtRatio >= 0.05) textLikeStrips++;
      totalStrips++;
    }

    return totalStrips > 0 ? textLikeStrips / totalStrips : 0;
  }

  function isUnsafeAnchorScore(metrics) {
    if (!metrics) return false;
    // Face/head check — driven by neural detection (skinRatio=1.0) or pixel skin ratio.
    //   skin:0.0000–0.03 → safe background, no face → SAFE
    //   skin:0.04+       → face/hand/arm present → UNSAFE
    if (metrics.skinRatio    > 0.04)  return true;
    if (metrics.topSkinRatio > 0.025) return true;
    // Text/graphics check — solid-background + contrast heuristic.
    //   textRatio > 0.10 means ≥10% of horizontal strips show solid bg + high-contrast text → UNSAFE
    //   New algorithm produces near-zero for natural content and 0.3–0.9 for real text overlays.
    if ((metrics.textRatio || 0) > 0.10) return true;
    return false;
  }

  function chooseAnchorFromScores(preferredAnchor, scores) {
    var preferred = scores[preferredAnchor] || scores['top-right'];
    var alternateKey = preferredAnchor === 'top-right' ? 'top-left' : 'top-right';
    var alternate = scores[alternateKey] || preferred;
    var preferredUnsafe = isUnsafeAnchorScore(preferred);
    var alternateUnsafe = isUnsafeAnchorScore(alternate);

    if (!preferredUnsafe && !alternateUnsafe) {
      if ((preferred.score - alternate.score) > 0.12) {
        return alternateKey;
      }
      return preferredAnchor;
    }
    if (preferredUnsafe && !alternateUnsafe) return alternateKey;
    if (!preferredUnsafe && alternateUnsafe) return preferredAnchor;
    return alternate.score < preferred.score ? alternateKey : preferredAnchor;
  }

  // ── Neural face detection (face-api.js / TinyFaceDetector) ─────────────────
  //
  // Replaces pixel-based skin analysis with a real neural network.
  // Models are loaded once at startup from client/models/.
  // Falls back silently to pixel analysis if models are unavailable.

  function initFaceApi() {
    var faceApiGlobal = (typeof faceapi !== 'undefined') ? faceapi
      : (typeof window !== 'undefined' && window.faceapi) ? window.faceapi
      : null;

    if (!faceApiGlobal) {
      log('[face-api] Library not loaded — using pixel-based analysis fallback.');
      return;
    }

    var modelsDir = extensionRoot ? path.join(extensionRoot, 'client', 'models') : '';
    if (!modelsDir || !fs.existsSync(modelsDir)) {
      log('[face-api] Models folder not found at "' + modelsDir + '" — pixel-based fallback active.');
      return;
    }

    var manifestFile = path.join(modelsDir, 'tiny_face_detector_model-weights_manifest.json');
    var binFile      = path.join(modelsDir, 'tiny_face_detector_model.bin');
    if (!fs.existsSync(manifestFile) || !fs.existsSync(binFile)) {
      log('[face-api] Model files missing in "' + modelsDir + '" — pixel-based fallback active.');
      return;
    }

    // Build a file:// URI that the Chromium/CEF XHR can reach.
    // Windows paths often contain spaces — encode them so the URL is valid.
    var modelsUri;
    if (process.platform === 'win32') {
      modelsUri = 'file:///' + modelsDir.replace(/\\/g, '/').replace(/ /g, '%20');
    } else {
      modelsUri = 'file://' + modelsDir.replace(/ /g, '%20');
    }

    log('[face-api] Loading TinyFaceDetector model from: ' + modelsUri);

    faceApiGlobal.nets.tinyFaceDetector.loadFromUri(modelsUri)
      .then(function () {
        faceApiReady = true;
        log('[face-api] ✓ TinyFaceDetector ready — neural face detection is active.');
      })
      .catch(function (err) {
        faceApiInitError = err && err.message ? err.message : String(err);
        log('[face-api] Model load failed: ' + faceApiInitError + ' — pixel-based fallback active.');
      });
  }

  // ── OCR text detection (OCRAD.js) ───────────────────────────────────────────
  //
  // OCRAD.js is a pure synchronous asm.js OCR engine — no Web Workers,
  // no SharedArrayBuffer, no WASM. Works in CEP's restricted Chromium environment.
  // Loaded via <script src="lib/ocrad.js"> in index.html.

  function initOCRAD() {
    var ocradFn = (typeof OCRAD !== 'undefined') ? OCRAD
      : (typeof window !== 'undefined' && window.OCRAD) ? window.OCRAD
      : null;
    if (!ocradFn) {
      log('[text-detect] OCRAD not available — pixel-based text detection active.');
      return;
    }
    ocradReady = true;
    log('[text-detect] ✓ OCRAD ready — OCR text detection active.');
  }

  // Synchronous OCR on one anchor region of a canvas.
  // Returns 0.9 if real text found (3+ consecutive letters), 0.0 otherwise.
  function detectTextWithOCRAD(canvas, anchorKey) {
    var ocradFn = (typeof OCRAD !== 'undefined') ? OCRAD
      : (typeof window !== 'undefined' && window.OCRAD) ? window.OCRAD
      : null;
    if (!ocradFn) return 0.0;

    var regionDefs = {
      'top-left':  { x: 0.15, y: 0.04, w: 0.28, h: 0.09 },
      'top-right': { x: 0.66, y: 0.02, w: 0.34, h: 0.07 }
    };
    var reg = regionDefs[anchorKey];
    if (!reg) return 0.0;

    var W = canvas.width, H = canvas.height;
    var rx = Math.floor(W * reg.x);
    var ry = Math.floor(H * reg.y);
    var rw = Math.min(Math.floor(W * reg.w), W - rx);
    var rh = Math.min(Math.floor(H * reg.h), H - ry);
    if (rw <= 0 || rh <= 0) return 0.0;

    // Upscale 2× — OCRAD reads small broadcast text better at higher resolution
    var scale = 2;
    var crop = document.createElement('canvas');
    crop.width  = rw * scale;
    crop.height = rh * scale;
    var cropCtx = crop.getContext('2d');
    cropCtx.drawImage(canvas, rx, ry, rw, rh, 0, 0, crop.width, crop.height);

    var N = crop.width * crop.height;
    var imgData = cropCtx.getImageData(0, 0, crop.width, crop.height);
    var d = imgData.data;

    // Build grayscale luminance array
    var lums = new Array(N);
    var pi, ti;
    for (pi = 0; pi < N; pi++) {
      lums[pi] = 0.299 * d[pi * 4] + 0.587 * d[pi * 4 + 1] + 0.114 * d[pi * 4 + 2];
    }

    // Skip flat regions (clear sky, solid colour) — no text can be there.
    // Variance < 300 means std-dev < ~17: too uniform to contain legible text.
    var mean = 0;
    for (pi = 0; pi < N; pi++) mean += lums[pi];
    mean /= N;
    var variance = 0;
    for (pi = 0; pi < N; pi++) {
      var dv = lums[pi] - mean;
      variance += dv * dv;
    }
    variance /= N;
    if (variance < 300) {
      return 0.0;
    }

    // Otsu's method — automatically finds the luminance threshold that best
    // separates text pixels from background pixels, regardless of text colour.
    var hist = new Array(256);
    for (ti = 0; ti < 256; ti++) hist[ti] = 0;
    for (pi = 0; pi < N; pi++) hist[Math.round(lums[pi])]++;
    var sumAll = 0;
    for (ti = 0; ti < 256; ti++) sumAll += ti * hist[ti];
    var sumB = 0, wB = 0, maxBetween = 0, otsuT = 128;
    for (ti = 0; ti < 256; ti++) {
      wB += hist[ti];
      if (!wB || wB === N) continue;
      var wF = N - wB;
      sumB += ti * hist[ti];
      var mB = sumB / wB;
      var mF = (sumAll - sumB) / wF;
      var between = wB * wF * (mB - mF) * (mB - mF);
      if (between > maxBetween) { maxBetween = between; otsuT = ti; }
    }

    // Binarize with a given threshold and run OCRAD.
    // invert=true → dark text on white bg (best for OCRAD); invert=false → white text on dark bg.
    var textRe = /[A-Za-z]{3,}/;
    function tryOCRADAt(threshold, invert, label) {
      for (var pp = 0; pp < N; pp++) {
        var v = (lums[pp] > threshold) ? 255 : 0;
        if (invert) v = 255 - v;
        d[pp * 4] = d[pp * 4 + 1] = d[pp * 4 + 2] = v;
        d[pp * 4 + 3] = 255;
      }
      cropCtx.putImageData(imgData, 0, 0);
      try {
        var raw = ocradFn(crop);
        var matched = textRe.test(raw);
        if (matched) {
          log('[ocr] ' + anchorKey + ' ✓ [' + label + '] → "' + raw.replace(/\n/g, ' ').trim().slice(0, 60) + '"');
        }
        return matched;
      } catch (e) { return false; }
    }

    // Pass 1 — Otsu inverted (dark text on white bg — OCRAD's preferred input)
    // Pass 2 — Otsu normal (catches dark-bg text that inverts poorly)
    // Pass 3 — Fixed t=200 inverted (isolates bright/white text; Otsu sits too low on blue sky)
    // Pass 4 — Fixed t=200 normal (fallback)
    return (tryOCRADAt(otsuT, true,  'otsu-inv') ||
            tryOCRADAt(otsuT, false, 'otsu')     ||
            tryOCRADAt(200,   true,  't200-inv')  ||
            tryOCRADAt(200,   false, 't200'))
           ? 0.9 : 0.0;
  }

  // Returns a Promise that resolves to { 'top-left': ratio, 'top-right': ratio }
  // ratio is 0.0 (no text) or 0.9 (text detected by OCRAD), or 0–1 from pixel fallback.
  // Never rejects.
  function detectTextInFrameAsync(canvas) {
    return new Promise(function (resolve) {
      if (ocradReady) {
        resolve({
          'top-left':  detectTextWithOCRAD(canvas, 'top-left'),
          'top-right': detectTextWithOCRAD(canvas, 'top-right')
        });
      } else {
        var ctx = canvas.getContext('2d');
        resolve({
          'top-left':  detectTextInRegion(ctx, canvas.width, canvas.height, 'top-left'),
          'top-right': detectTextInRegion(ctx, canvas.width, canvas.height, 'top-right')
        });
      }
    });
  }

  function runFaceApiAnalysis(canvas, imgElement, preferredAnchor, callback) {
    var w = canvas.width;
    var h = canvas.height;

    // Regions match the actual slide badge footprint measured from broadcast frames:
    //   top-left : x 15–43%, y 4–13%
    //   top-right: x 66–100%, y 2–9%
    var regionDefs = {
      'top-left':  { x1: 0.15, y1: 0.04, x2: 0.43, y2: 0.13 },
      'top-right': { x1: 0.66, y1: 0.02, x2: 1.00, y2: 0.09 }
    };

    var faceApiGlobal;
    try {
      faceApiGlobal = (typeof faceapi !== 'undefined') ? faceapi : window.faceapi;
    } catch(e) { faceApiGlobal = null; }

    if (!faceApiGlobal) {
      log('[face-api] ⚠ faceapi global not accessible inside runFaceApiAnalysis — pixel fallback.');
      var ctx0 = canvas.getContext('2d');
      var ps0 = { 'top-left': scoreAnchorRegion(ctx0, w, h, 'top-left'), 'top-right': scoreAnchorRegion(ctx0, w, h, 'top-right') };
      ps0['top-left'].textRatio  = detectTextInRegion(ctx0, w, h, 'top-left');
      ps0['top-right'].textRatio = detectTextInRegion(ctx0, w, h, 'top-right');
      return callback(null, { resolvedAnchor: chooseAnchorFromScores(preferredAnchor, ps0), scores: ps0, allUnsafe: false });
    }

    var detectorOptions;
    try {
      detectorOptions = new faceApiGlobal.TinyFaceDetectorOptions({ scoreThreshold: 0.4, inputSize: 320 });
    } catch(e) {
      log('[face-api] ⚠ TinyFaceDetectorOptions constructor error: ' + e.message + ' — pixel fallback.');
      var ctx1 = canvas.getContext('2d');
      var ps1 = { 'top-left': scoreAnchorRegion(ctx1, w, h, 'top-left'), 'top-right': scoreAnchorRegion(ctx1, w, h, 'top-right') };
      ps1['top-left'].textRatio  = detectTextInRegion(ctx1, w, h, 'top-left');
      ps1['top-right'].textRatio = detectTextInRegion(ctx1, w, h, 'top-right');
      return callback(null, { resolvedAnchor: chooseAnchorFromScores(preferredAnchor, ps1), scores: ps1, allUnsafe: false });
    }

    // Use the img element directly for face detection — face-api handles HTMLImageElement
    // natively and avoids any canvas-context issues.
    var faceInput = (imgElement && imgElement.naturalWidth) ? imgElement : canvas;
    // Image dimensions for bbox normalisation
    var iw = (faceInput === imgElement && imgElement.naturalWidth) ? imgElement.naturalWidth  : w;
    var ih = (faceInput === imgElement && imgElement.naturalHeight) ? imgElement.naturalHeight : h;

    // Run face detection (face-api) AND text detection (Tesseract) in parallel.
    // detectTextInFrameAsync never rejects — falls back to pixel on any error.
    Promise.all([
      faceApiGlobal.detectAllFaces(faceInput, detectorOptions),
      detectTextInFrameAsync(canvas)
    ])
      .then(function (results) {
        var detections  = results[0];
        var textRatioMap = results[1];  // { 'top-left': 0|0.9, 'top-right': 0|0.9 }
        var scores = {};
        var anchorKeys = ['top-left', 'top-right'];

        anchorKeys.forEach(function (anchorKey) {
          var reg = regionDefs[anchorKey];
          var faceDetected = false;
          var maxConfidence = 0;
          var textRatio = textRatioMap[anchorKey] || 0;

          detections.forEach(function (det) {
            // Normalize detected bounding box to 0–1 fractions of image size
            var bx1 = det.box.x / iw;
            var by1 = det.box.y / ih;
            var bx2 = (det.box.x + det.box.width)  / iw;
            var by2 = (det.box.y + det.box.height) / ih;

            // AABB intersection with the slide placement region
            if (bx2 > reg.x1 && bx1 < reg.x2 && by2 > reg.y1 && by1 < reg.y2) {
              faceDetected = true;
              if (det.score > maxConfidence) maxConfidence = det.score;
            }
          });

          // Combine face and text signals into a unified "skinRatio-equivalent"
          // that drives all downstream safety gates (isUnsafeAnchorScore etc.).
          //   faceDetected → skinRatio 0.9+  (always unsafe)
          //   text only    → skinRatio = 0 but textRatio checked separately
          //   both clear   → skinRatio 0.0, textRatio ~0 (safe)
          var skinRatioEquiv = faceDetected ? Math.max(0.9, maxConfidence) : 0.0;
          scores[anchorKey] = {
            anchor:               anchorKey,
            score:                (faceDetected || textRatio > 0.15) ? 1.0 : 0.0,
            skinRatio:            skinRatioEquiv,
            topSkinRatio:         skinRatioEquiv,
            darkRatio:            0,
            lowerOccupiedRatio:   0,
            centralOccupiedRatio: 0,
            clearanceRatio:       0,
            belowBandRatio:       0,
            edgeRatio:            0,
            textureRatio:         0,
            textRatio:            textRatio,
            faceDetected:         faceDetected,
            faceCount:            detections.length,
            maxConfidence:        maxConfidence
          };
        });

        var resolvedAnchor = chooseAnchorFromScores(preferredAnchor, scores);
        callback(null, {
          resolvedAnchor: resolvedAnchor,
          reason: resolvedAnchor === preferredAnchor ? 'preferred-safe' : 'switched-for-head-face-avoidance',
          scores: scores,
          allUnsafe: false,
          faceCount: detections.length
        });
      })
      .catch(function (detErr) {
        // face-api detection error for this frame → log it, then fall back to pixel analysis
        var errMsg = detErr && detErr.message ? detErr.message : String(detErr);
        log('[face-api] ⚠ detectAllFaces error: ' + errMsg + ' — pixel fallback used for this frame.');
        var ctx2 = canvas.getContext('2d');
        var pixelScores = {
          'top-left':  scoreAnchorRegion(ctx2, canvas.width, canvas.height, 'top-left'),
          'top-right': scoreAnchorRegion(ctx2, canvas.width, canvas.height, 'top-right')
        };
        // Add text detection to the pixel-fallback scores too
        pixelScores['top-left'].textRatio  = detectTextInRegion(ctx2, canvas.width, canvas.height, 'top-left');
        pixelScores['top-right'].textRatio = detectTextInRegion(ctx2, canvas.width, canvas.height, 'top-right');
        var resolvedAnchor = chooseAnchorFromScores(preferredAnchor, pixelScores);
        callback(null, {
          resolvedAnchor: resolvedAnchor,
          reason: resolvedAnchor === preferredAnchor ? 'preferred-safe' : 'switched-for-head-face-avoidance',
          scores: pixelScores,
          allUnsafe: false
        });
      });
  }

  function analyzeVisibleFrameAnchor(framePath, preferredAnchor, callback) {
    loadImageFromFile(framePath, function (err, img) {
      if (err) {
        callback(null, {
          resolvedAnchor: preferredAnchor,
          reason: 'frame-unavailable',
          scores: null
        });
        return;
      }

      var canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth || img.width;
      canvas.height = img.naturalHeight || img.height;
      var ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

      if (faceApiReady) {
        // Neural network face detection + pixel text detection (async, Promise-based)
        runFaceApiAnalysis(canvas, img, preferredAnchor, callback);
      } else {
        // Pixel-based skin analysis + text detection fallback
        var scores = {
          'top-left':  scoreAnchorRegion(ctx, canvas.width, canvas.height, 'top-left'),
          'top-right': scoreAnchorRegion(ctx, canvas.width, canvas.height, 'top-right')
        };
        scores['top-left'].textRatio  = detectTextInRegion(ctx, canvas.width, canvas.height, 'top-left');
        scores['top-right'].textRatio = detectTextInRegion(ctx, canvas.width, canvas.height, 'top-right');
        var resolvedAnchor = chooseAnchorFromScores(preferredAnchor, scores);
        callback(null, {
          resolvedAnchor: resolvedAnchor,
          reason: resolvedAnchor === preferredAnchor ? 'preferred-safe' : 'switched-for-head-face-avoidance',
          scores: scores,
          allUnsafe: false
        });
      }
    });
  }

  function mergeAnchorScores(preferredAnchor, analyses) {
    var merged = {
      'top-left': {
        anchor: 'top-left',
        score: 0,
        skinRatio: 0,
        topSkinRatio: 0,
        darkRatio: 0,
        lowerOccupiedRatio: 0,
        centralOccupiedRatio: 0,
        clearanceRatio: 0,
        belowBandRatio: 0,
        edgeRatio: 0,
        textureRatio: 0,
        textRatio: 0,
        faceDetected: false,
        maxConfidence: 0,
        faceCount: 0
      },
      'top-right': {
        anchor: 'top-right',
        score: 0,
        skinRatio: 0,
        topSkinRatio: 0,
        darkRatio: 0,
        lowerOccupiedRatio: 0,
        centralOccupiedRatio: 0,
        clearanceRatio: 0,
        belowBandRatio: 0,
        edgeRatio: 0,
        textureRatio: 0,
        textRatio: 0,
        faceDetected: false,
        maxConfidence: 0,
        faceCount: 0
      }
    };

    analyses.forEach(function (analysis) {
      if (!analysis || !analysis.scores) return;
      ['top-left', 'top-right'].forEach(function (anchorKey) {
        var source = analysis.scores[anchorKey];
        var target = merged[anchorKey];
        if (!source) return;
        target.score = Math.max(target.score, source.score || 0);
        target.skinRatio = Math.max(target.skinRatio, source.skinRatio || 0);
        target.topSkinRatio = Math.max(target.topSkinRatio, source.topSkinRatio || 0);
        target.darkRatio = Math.max(target.darkRatio, source.darkRatio || 0);
        target.lowerOccupiedRatio = Math.max(target.lowerOccupiedRatio, source.lowerOccupiedRatio || 0);
        target.centralOccupiedRatio = Math.max(target.centralOccupiedRatio, source.centralOccupiedRatio || 0);
        target.clearanceRatio = Math.max(target.clearanceRatio, source.clearanceRatio || 0);
        target.belowBandRatio = Math.max(target.belowBandRatio, source.belowBandRatio || 0);
        target.edgeRatio = Math.max(target.edgeRatio, source.edgeRatio || 0);
        target.textureRatio = Math.max(target.textureRatio, source.textureRatio || 0);
        target.textRatio = Math.max(target.textRatio, source.textRatio || 0);
        // Preserve NN-specific fields — if ANY frame detected a face, mark the whole slot unsafe
        if (source.faceDetected) target.faceDetected = true;
        if ((source.maxConfidence || 0) > target.maxConfidence) target.maxConfidence = source.maxConfidence;
        if ((source.faceCount || 0) > target.faceCount) target.faceCount = source.faceCount;
      });
    });

    return {
      resolvedAnchor: chooseAnchorFromScores(preferredAnchor, merged),
      scores: merged,
      allUnsafe: (isUnsafeAnchorScore(merged['top-left']) && isUnsafeAnchorScore(merged['top-right']))
    };
  }

  function analyzeVisibleFrameSet(framePaths, preferredAnchor, callback) {
    var paths = Array.isArray(framePaths) ? framePaths.filter(Boolean) : [];
    if (!paths.length) {
      callback(null, {
        resolvedAnchor: preferredAnchor,
        reason: 'frame-unavailable',
        scores: null
      });
      return;
    }

    var analyses = [];
    var index = 0;
    function analyzeNextFrame() {
      if (index >= paths.length) {
        var merged = mergeAnchorScores(preferredAnchor, analyses);
        callback(null, {
          resolvedAnchor: merged.resolvedAnchor,
          reason: merged.resolvedAnchor === preferredAnchor ? 'preferred-safe-across-samples' : 'switched-for-head-face-avoidance',
          scores: merged.scores,
          allUnsafe: !!merged.allUnsafe
        });
        return;
      }

      analyzeVisibleFrameAnchor(paths[index], preferredAnchor, function (err, analysis) {
        if (analysis && analysis.scores) {
          analyses.push(analysis);
        }
        index++;
        analyzeNextFrame();
      });
    }

    analyzeNextFrame();
  }

  function resolveAnchorsForPlacementPlan(placementPlan, preferredAnchor, enabled, callback) {
    var placements = placementPlan && Array.isArray(placementPlan.placements) ? placementPlan.placements : [];
    var summary = {
      analyzedCount: 0,
      switchedCount: 0,
      missingFrameCount: 0,
      unsafeCount: 0
    };

    if (!enabled || !placements.length) {
      placements.forEach(function (placement) {
        placement.resolvedAnchor = preferredAnchor;
      });
      callback(null, summary);
      return;
    }

    var index = 0;
    function analyzeNext() {
      if (index >= placements.length) {
        callback(null, summary);
        return;
      }

      var placement = placements[index];
      var framePaths = Array.isArray(placement.framePaths) ? placement.framePaths : (placement.framePath ? [placement.framePath] : []);
      if (!framePaths.length) {
        placement.resolvedAnchor = preferredAnchor;
        placement.anchorReason = 'frame-unavailable';
        placement.allUnsafe = false;
        summary.missingFrameCount++;
        index++;
        analyzeNext();
        return;
      }

      analyzeVisibleFrameSet(framePaths, preferredAnchor, function (err, analysis) {
        placement.resolvedAnchor = (analysis && analysis.resolvedAnchor) ? analysis.resolvedAnchor : preferredAnchor;
        placement.anchorReason = analysis && analysis.reason ? analysis.reason : 'frame-unavailable';
        placement.allUnsafe = !!(analysis && analysis.allUnsafe);
        if (!analysis || analysis.reason === 'frame-unavailable') {
          summary.missingFrameCount++;
        } else {
          summary.analyzedCount++;
        }
        if (placement.allUnsafe) {
          summary.unsafeCount++;
        }
        if (placement.resolvedAnchor !== preferredAnchor) {
          summary.switchedCount++;
        }
        index++;
        analyzeNext();
      });
    }

    analyzeNext();
  }

  function overlapsBlockedRange(startSeconds, durationSeconds, blockedRanges) {
    if (!Array.isArray(blockedRanges) || !blockedRanges.length) return false;
    var endSeconds = startSeconds + Math.max(0, durationSeconds || 0);
    return blockedRanges.some(function (range) {
      return range && typeof range.start === 'number' && typeof range.end === 'number' && endSeconds > range.start && startSeconds < range.end;
    });
  }

  function clonePlacementPlan(plan) {
    return JSON.parse(JSON.stringify(plan || {}));
  }

  function nextAvailableStart(startSeconds, durationSeconds, blockedRanges) {
    var candidate = Math.max(0, startSeconds || 0);
    var moved = true;
    while (moved) {
      moved = false;
      if (Array.isArray(blockedRanges)) {
        for (var i = 0; i < blockedRanges.length; i++) {
          var range = blockedRanges[i];
          if (!range || typeof range.start !== 'number' || typeof range.end !== 'number') continue;
          if ((candidate + durationSeconds) > range.start && candidate < range.end) {
            candidate = range.end;
            moved = true;
            break;
          }
        }
      }
    }
    return candidate;
  }

  function previewSinglePlacement(options, placement, startSeconds, callback) {
    var analysisDir = createFaceAnalysisTempDir();
    callJsx('newPeaceMakerPreviewSinglePlacement', {
      analysisDir: analysisDir,
      startSeconds: startSeconds,
      clipDurationSeconds: placement.clipDurationSeconds || 1,
      categoryName: placement.categoryName,
      language: placement.language,
      placementIndex: placement.placementIndex || 0
    }, function (result) {
      // NOTE: do NOT cleanupDir here — the frame files must exist on disk
      // while analyzeVisibleFrameSet reads them. Clean up AFTER analysis.
      var parsed;
      try {
        parsed = JSON.parse(result);
      } catch (e) {
        cleanupDir(analysisDir);
        callback(new Error('Single placement preview failed: ' + result));
        return;
      }
      if (!parsed.ok || !parsed.placementPreview) {
        cleanupDir(analysisDir);
        callback(new Error(parsed.error || 'Single placement preview failed.'));
        return;
      }

      var preview = parsed.placementPreview;
      var framePaths = preview.framePaths || [];
      var exportedCount = framePaths.filter(Boolean).length;

      // Diagnostic: report frame export status
      var sampleTimes = preview.sampleTimes || [];
      var sampleLabel = sampleTimes.length
        ? ' [' + sampleTimes.map(function (t) { return secToMS(t); }).join(', ') + ']'
        : '';
      if (exportedCount === 0) {
        log('  ⚠ No frames exported at ' + secToMS(startSeconds) + ' — frame analysis skipped, using preferred anchor.');
        log('  Export errors: ' + (parsed.exportErrors && parsed.exportErrors.length ? parsed.exportErrors.join(' | ') : 'none reported'));
      } else {
        log('  ✓ ' + exportedCount + ' frame(s) exported' + sampleLabel + ' — analyzing for face/head...');
      }

      analyzeVisibleFrameSet(framePaths, options.slideAnchor, function (err, analysis) {
        cleanupDir(analysisDir); // clean up AFTER images have been read

        // Diagnostic: log scores so we can see what detection found
        if (analysis && analysis.scores) {
          var tl = analysis.scores['top-left'];
          var tr = analysis.scores['top-right'];
          var usingNN = faceApiReady && tl && typeof tl.faceDetected !== 'undefined';
          var tlInfo = tl
            ? (usingNN
                ? (tl.faceDetected ? 'FACE(conf:' + (tl.maxConfidence || 0).toFixed(2) + ')' : 'clear') +
                  ((tl.textRatio || 0) > 0.15 ? '+TEXT' : '')
                : tl.score.toFixed(3) + ' skin:' + tl.skinRatio.toFixed(4))
            : 'n/a';
          var trInfo = tr
            ? (usingNN
                ? (tr.faceDetected ? 'FACE(conf:' + (tr.maxConfidence || 0).toFixed(2) + ')' : 'clear') +
                  ((tr.textRatio || 0) > 0.15 ? '+TEXT' : '')
                : tr.score.toFixed(3) + ' skin:' + tr.skinRatio.toFixed(4))
            : 'n/a';
          var tlTextInfo = (!usingNN && tl && (tl.textRatio || 0) > 0) ? ' txt:' + (tl.textRatio).toFixed(2) : '';
          var trTextInfo = (!usingNN && tr && (tr.textRatio || 0) > 0) ? ' txt:' + (tr.textRatio).toFixed(2) : '';
          log('  ' + (usingNN ? '[NN]' : '[px]') + ' top-left: ' + tlInfo + tlTextInfo +
              ' | top-right: ' + trInfo + trTextInfo +
              (usingNN ? ' faces-in-frame:' + (analysis.faceCount || 0) : '') +
              ' | chosen: ' + (analysis.resolvedAnchor || '?') +
              (analysis.allUnsafe ? ' ⚠ BOTH UNSAFE' : ''));
        }

        callback(null, {
          startSeconds: startSeconds,
          clipDurationSeconds: preview.clipDurationSeconds || placement.clipDurationSeconds || 1,
          resolvedAnchor: analysis && analysis.resolvedAnchor ? analysis.resolvedAnchor : options.slideAnchor,
          allUnsafe: !!(analysis && analysis.allUnsafe),
          scores: analysis ? analysis.scores : null
        });
      });
    });
  }

  function rerenderPlacementPlan(options, placementPlan, callback) {
    var analysisDir = '';
    if (options.avoidFaces) {
      analysisDir = createFaceAnalysisTempDir();
    }

    callJsx('newPeaceMakerPreviewPlacementFrames', {
      batches: options.batches,
      targetTrack: options.targetTrack,
      ignoreV1: options.ignoreV1,
      analysisDir: analysisDir,
      exportFrames: !!options.avoidFaces,
      placementPlan: placementPlan
    }, function (previewResult) {
      cleanupDir(analysisDir);
      var parsed;
      try {
        parsed = JSON.parse(previewResult);
      } catch (e) {
        callback(new Error('Preview failed: ' + previewResult));
        return;
      }
      if (!parsed.ok) {
        callback(new Error(parsed.error || 'Preview failed.'));
        return;
      }
      resolveAnchorsForPlacementPlan(parsed.placementPlan, options.slideAnchor, options.avoidFaces, function (err, summary) {
        callback(null, parsed, summary);
      });
    });
  }

  function moveUnsafePlacements(options, placementPlan, callback) {
    var placements = placementPlan && Array.isArray(placementPlan.placements) ? placementPlan.placements : [];
    var placementWindowEnd = placementPlan && typeof placementPlan.placementWindowEndSeconds === 'number' ? placementPlan.placementWindowEndSeconds : 0;
    var blockedRanges = placementPlan && Array.isArray(placementPlan.blockedV1Ranges) ? placementPlan.blockedV1Ranges : [];
    var movedCount = 0;
    var index = 0;

    function tryNextPlacement() {
      while (index < placements.length && !placements[index].allUnsafe) {
        index++;
      }
      if (index >= placements.length) {
        callback(null, placementPlan, movedCount);
        return;
      }

      var placement = placements[index];
      var originalStart = placement.startSeconds || 0;
      var duration = placement.clipDurationSeconds || 1;
      var maxStart = placementWindowEnd > 0 ? Math.max(originalStart, placementWindowEnd - duration) : (originalStart + 12);
      var attempts = [];
      var step = 0.25;
      for (var d = 1; d <= 48; d++) {
        attempts.push(originalStart + (step * d));
        attempts.push(Math.max(0, originalStart - (step * d)));
      }
      attempts = attempts.filter(function (candidate) {
        return candidate >= 0 && candidate <= maxStart && !overlapsBlockedRange(candidate, duration, blockedRanges);
      });

      function tryCandidateAt(attemptIndex) {
        if (attemptIndex >= attempts.length) {
          index++;
          tryNextPlacement();
          return;
        }

        var candidateStart = attempts[attemptIndex];
        var candidatePlan = clonePlacementPlan(placementPlan);
        candidatePlan.placements[index].startSeconds = candidateStart;
        rerenderPlacementPlan(options, candidatePlan, function (err, previewParsed) {
          if (err || !previewParsed || !previewParsed.placementPlan || !previewParsed.placementPlan.placements[index]) {
            tryCandidateAt(attemptIndex + 1);
            return;
          }
          var candidatePlacement = previewParsed.placementPlan.placements[index];
          if (!candidatePlacement.allUnsafe) {
            placementPlan = previewParsed.placementPlan;
            placements = placementPlan.placements;
            movedCount++;
            index++;
            tryNextPlacement();
            return;
          }
          tryCandidateAt(attemptIndex + 1);
        });
      }

      tryCandidateAt(0);
    }

    tryNextPlacement();
  }

  function buildSafePlacementPlan(options, basePlan, callback) {
    // Rules:
    // 1. Each category gets an equal zone of [windowStart, windowEnd].
    //    Zones are HARD search limits — slides never jump into the next category's zone
    //    just because the current zone is face/text-saturated.  This prevents a single
    //    "busy" zone from cascading all remaining slides into the end of the timeline.
    // 2. Face-avoidance searches ±3s (coarse) then ±1s (fine) from the ideal slot,
    //    interleaving forward and backward steps.
    // 3. cursor tracks end of last placed slide — next slide NEVER starts before cursor
    //    (prevents Premiere overwriteClip from trimming an already-placed slide).
    // 4. Hard deadline per slide = min(catZone.end, windowEnd) − duration.
    // 5. If cursor > deadline (zone exhausted by earlier slides), overflow slides are
    //    appended SEQUENTIALLY at cursor — never stacked at the same timecode.
    // 6. If no safe spot exists within the zone → use the least-bad position found.
    // 7. Use actual clip duration — never force-clamp to a minimum.
    var placementPlan = clonePlacementPlan(basePlan);
    var placements    = placementPlan && Array.isArray(placementPlan.placements) ? placementPlan.placements : [];
    var blockedRanges = placementPlan && Array.isArray(placementPlan.blockedV1Ranges) ? placementPlan.blockedV1Ranges : [];

    var windowStart = typeof placementPlan.placementWindowStartSeconds === 'number' ? placementPlan.placementWindowStartSeconds : 0;
    // windowEnd: prefer in/out value, then fall back to max clip end, then 0 (handled below)
    var windowEnd   = 0;
    if (typeof placementPlan.placementWindowEndSeconds === 'number' && placementPlan.placementWindowEndSeconds > 0) {
      windowEnd = placementPlan.placementWindowEndSeconds;
    } else if (typeof placementPlan.usedTimelineLengthSeconds === 'number' && placementPlan.usedTimelineLengthSeconds > 0) {
      windowEnd = placementPlan.usedTimelineLengthSeconds;
    }
    if (windowEnd <= windowStart) {
      callback(new Error('Timeline window has zero usable length (windowEnd=' + windowEnd + ', windowStart=' + windowStart + ').'));
      return;
    }

    var skipStep = 3;   // seconds per search step
    var slideGap = 0.1; // minimum gap between consecutive slides (prevents exact overlap)
    var movedCount = 0;
    var unsafeFallbackCount = 0;
    var index  = 0;
    var cursor = windowStart; // end of last placed slide — next slide MUST start here or later

    // ── Per-category time zones ──────────────────────────────────────────────
    // Categories are already in fixed order (categoryOrder array governs batches).
    // We give each category an equal slice of the total window so their slides
    // can never mix with adjacent categories even if face-avoidance shifts them.
    var categoryNames = [];
    var seenCats = {};
    for (var ci = 0; ci < placements.length; ci++) {
      var cname = placements[ci].categoryName || ('_cat' + ci);
      if (!seenCats[cname]) { seenCats[cname] = true; categoryNames.push(cname); }
    }
    var catCount = categoryNames.length || 1;

    // Map a "compressed" duration (V1-excluded time) back to a real timeline position.
    // Walks forward from windowStart, skipping blocked V1 ranges, until the
    // requested amount of free time has elapsed.  Falls back to raw offset when
    // ignoreV1 is off (blockedRanges is empty).
    function mapCompressedToReal(compressedSec, bRanges, wStart, wEnd) {
      if (!bRanges || !bRanges.length) return Math.min(wStart + compressedSec, wEnd);
      var sorted = bRanges.slice().sort(function (a, b) { return a.start - b.start; });
      var remaining = compressedSec;
      var t = wStart;
      for (var ri = 0; ri < sorted.length; ri++) {
        var bStart = Math.max(sorted[ri].start, wStart);
        var bEnd   = Math.min(sorted[ri].end,   wEnd);
        if (bStart >= wEnd) break;
        var freeInGap = Math.max(0, bStart - t);
        if (remaining <= freeInGap) return t + remaining;
        remaining -= freeInGap;
        t = Math.max(t, bEnd);
      }
      return Math.min(t + remaining, wEnd);
    }

    // Use V1-excluded ("usable") time to define fair zone boundaries so that
    // categories whose raw time slot is heavily blocked by V1 get the same
    // effective free-time budget as other categories.
    var usableLen = (placementPlan && placementPlan.usableTimelineLengthSeconds)
                  ? placementPlan.usableTimelineLengthSeconds
                  : (windowEnd - windowStart);
    var usablePerCat = usableLen / catCount;
    var catZones = {};
    for (var cz = 0; cz < categoryNames.length; cz++) {
      catZones[categoryNames[cz]] = {
        start: mapCompressedToReal(cz       * usablePerCat, blockedRanges, windowStart, windowEnd),
        end:   mapCompressedToReal((cz + 1) * usablePerCat, blockedRanges, windowStart, windowEnd)
      };
    }
    var zoneLen = usablePerCat; // for display
    log('Timeline window: ' + secToMS(windowStart) + ' – ' + secToMS(windowEnd) + '  |  ' +
        categoryNames.length + ' categories × preferred zone ' + zoneLen.toFixed(1) + 's usable  (hard limit: full window)');

    // Count slides per category for elastic spacing and skip+retry
    var catSlideCounts   = {};
    var catPlacedCounts  = {};
    var catPlacedIntervals = {}; // [{start,end}] per category — for gap-finding on retry
    var deferredSlides   = {};  // slides skipped due to full face-block; retried at cat end

    for (var ci = 0; ci < placements.length; ci++) {
      var cn = placements[ci].categoryName || ('_cat' + ci);
      catSlideCounts[cn]   = (catSlideCounts[cn]   || 0) + 1;
      catPlacedCounts[cn]  = 0;
      catPlacedIntervals[cn] = [];
      deferredSlides[cn]   = [];
    }

    // ── Position cache ────────────────────────────────────────────────────────
    // Keyed by Math.round(startSeconds). Stores face/text analysis so later
    // slides skip frame export for positions already checked (1s margin).
    var posCache = {};

    function previewCached(opts, pmt, t, cb) {
      var key = Math.round(t);
      if (posCache[key] !== undefined) {
        var c = posCache[key];
        // Reconstruct a minimal preview object from cached data
        cb(null, c ? {
          startSeconds:        t,
          clipDurationSeconds: pmt.clipDurationSeconds || 9,
          resolvedAnchor:      c.resolvedAnchor,
          allUnsafe:           c.allUnsafe,
          scores:              c.scores,
          framePaths:          []
        } : null);
        return;
      }
      previewSinglePlacement(opts, pmt, t, function (err, result) {
        posCache[key] = (err || !result) ? null : {
          resolvedAnchor: result.resolvedAnchor,
          allUnsafe:      result.allUnsafe,
          scores:         result.scores
        };
        cb(err, result);
      });
    }

    // ── Gap finder ────────────────────────────────────────────────────────────
    // Returns candidate start-times that fit 'dur' seconds inside gaps between
    // already-placed intervals in the given zone.
    function gapCandidates(zoneSt, zoneEd, intervals, dur) {
      var sorted = intervals.slice().sort(function (a, b) { return a.start - b.start; });
      var out = [];
      var t = zoneSt;
      for (var ii = 0; ii < sorted.length; ii++) {
        var gEnd = sorted[ii].start - slideGap;
        var s = t;
        while (s + dur <= gEnd) { out.push(Math.round(s * 10) / 10); s += skipStep; }
        t = sorted[ii].end + slideGap;
      }
      var s2 = t;
      while (s2 + dur <= zoneEd) { out.push(Math.round(s2 * 10) / 10); s2 += skipStep; }
      return out;
    }

    // ── Deferred-slide retry ──────────────────────────────────────────────────
    // Called when a category finishes (or all slides are done). Tries to fit
    // each deferred slide into a gap between already-placed slides using the
    // position cache (no repeat frame exports).
    function retryDeferred(catName, done) {
      var deferred = deferredSlides[catName];
      if (!deferred || !deferred.length) { done(); return; }
      var zone = catZones[catName] || { start: windowStart, end: windowEnd };

      function retryOne(di) {
        if (di >= deferred.length) { done(); return; }
        var pmt  = deferred[di];
        var dur  = pmt.clipDurationSeconds || 9;
        var gaps = gapCandidates(zone.start, zone.end - dur,
                                 catPlacedIntervals[catName], dur);
        // Sort gaps by proximity to ideal elastic position within zone
        var totalInCat  = catSlideCounts[catName]  || 1;
        var placedInCat = catPlacedCounts[catName] || 0;
        var idealTarget = zone.start + (zone.end - zone.start) * (placedInCat + 1) / totalInCat;
        gaps.sort(function (a, b) { return Math.abs(a - idealTarget) - Math.abs(b - idealTarget); });

        if (!gaps.length) {
          // No gap at all — fallback: place sequentially at cursor
          pmt.startSeconds        = Math.max(cursor, zone.end);
          pmt.clipDurationSeconds = dur;
          pmt.resolvedAnchor      = options.slideAnchor;
          pmt.allUnsafe           = true;
          unsafeFallbackCount++;
          cursor = pmt.startSeconds + dur + slideGap;
          log('Slide (deferred) [' + catName + ']: no gap — sequential at ' + secToMS(pmt.startSeconds) + '.');
          retryOne(di + 1);
          return;
        }

        var gi = 0;
        function tryGap() {
          if (gi >= gaps.length) {
            // All gap positions face-blocked — use first gap position as unsafe fallback
            pmt.startSeconds        = gaps[0];
            pmt.clipDurationSeconds = dur;
            pmt.resolvedAnchor      = options.slideAnchor;
            pmt.allUnsafe           = true;
            unsafeFallbackCount++;
            catPlacedIntervals[catName].push({ start: gaps[0], end: gaps[0] + dur });
            catPlacedCounts[catName] = (catPlacedCounts[catName] || 0) + 1;
            log('Slide (deferred) [' + catName + ']: all gap positions unsafe — placed at ' + secToMS(gaps[0]) + '.');
            retryOne(di + 1);
            return;
          }
          var gt = gaps[gi++];
          previewCached(options, pmt, gt, function (err, preview) {
            if (!err && preview && !preview.allUnsafe) {
              pmt.startSeconds        = gt;
              pmt.clipDurationSeconds = dur;
              pmt.resolvedAnchor      = preview.resolvedAnchor || options.slideAnchor;
              pmt.allUnsafe           = false;
              catPlacedIntervals[catName].push({ start: gt, end: gt + dur });
              catPlacedCounts[catName] = (catPlacedCounts[catName] || 0) + 1;
              movedCount++;
              log('Slide (deferred) [' + catName + ']: ✓ placed in gap at ' + secToMS(gt) +
                  '. Corner: ' + pmt.resolvedAnchor + '.');
              retryOne(di + 1);
            } else {
              tryGap();
            }
          });
        }
        tryGap();
      }
      retryOne(0);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────
    function combinedScore(preview) {
      if (!preview || !preview.scores) return 9999;
      var L = preview.scores['top-left']  ? preview.scores['top-left'].score  : 9999;
      var R = preview.scores['top-right'] ? preview.scores['top-right'].score : 9999;
      return L + R;
    }

    // ── Main placement loop ──────────────────────────────────────────────────
    var prevCatName = null; // track category changes to trigger deferred retry

    function placeNext() {
      if (index >= placements.length) {
        // Retry any deferred slides for the last category before finishing
        if (prevCatName && deferredSlides[prevCatName] && deferredSlides[prevCatName].length) {
          retryDeferred(prevCatName, function () {
            callback(null, placementPlan, { movedCount: movedCount, unsafeFallbackCount: unsafeFallbackCount });
          });
        } else {
          callback(null, placementPlan, { movedCount: movedCount, unsafeFallbackCount: unsafeFallbackCount });
        }
        return;
      }

      var placement = placements[index];
      var duration  = placement.clipDurationSeconds > 0 ? placement.clipDurationSeconds : 9;
      var catName   = placement.categoryName || ('_cat' + index);

      // Category changed → retry deferred slides for the previous category first
      if (prevCatName && catName !== prevCatName &&
          deferredSlides[prevCatName] && deferredSlides[prevCatName].length) {
        retryDeferred(prevCatName, function () {
          prevCatName = catName;
          placeNext();
        });
        return;
      }
      prevCatName = catName;

      var zone      = catZones[catName] || { start: windowStart, end: windowEnd };
      var hardDeadline = Math.min(zone.end, windowEnd) - duration;

      // Elastic spacing with a special case: the FIRST slide in each category
      // (the English slide) anchors at the zone start — no leading gap.
      // Remaining slides divide the leftover space equally.
      var totalInCat    = catSlideCounts[catName]  || 1;
      var placedInCat   = catPlacedCounts[catName] || 0;
      var remainingInCat = totalInCat - placedInCat;
      var isFirstInCat  = placedInCat === 0;
      var originalStart;
      if (isFirstInCat) {
        originalStart = Math.max(cursor, zone.start);   // English slide: start of zone
      } else {
        var remainingZone = Math.max(0, hardDeadline - cursor);
        originalStart = cursor + (remainingInCat > 0 ? remainingZone / remainingInCat : 0);
      }

      // Cursor overflow: when face-avoidance has pushed the cursor past the zone's hard
      // deadline, there is no more room inside the zone.  Place remaining slides of this
      // category SEQUENTIALLY at cursor (never at hardDeadline, which would stack them
      // on top of each other and cause Premiere to trim previously placed clips).
      // cursor is allowed to advance past zone.end into the next category's zone — the
      // next category's search simply starts from wherever cursor ends up.
      if (cursor > hardDeadline) {
        // Zone space is gone. Before deferring or going sequential, check if any
        // gap between already-placed slides can fit this slide.
        var moreInCat = placements.slice(index + 1).some(function (p) {
          return (p.categoryName || '') === catName;
        });
        var freeGaps = gapCandidates(zone.start, zone.end - duration, catPlacedIntervals[catName] || [], duration);
        if (freeGaps.length > 0) {
          var fgIdx = 0;
          (function tryFreeGap() {
            if (fgIdx >= freeGaps.length) {
              // All gap positions face-blocked — fall through to defer or sequential
              if (moreInCat) {
                log('Slide ' + (index + 1) + ' [' + catName + ']: zone exhausted, gaps all blocked — deferring.');
                deferredSlides[catName].push(placement);
                index++;
                placeNext();
                return;
              }
              var emergencyStart = Math.max(0, cursor);
              log('Slide ' + (index + 1) + ' [' + catName + ']: zone exhausted, no clear gap — sequentially at ' +
                  secToMS(emergencyStart) + '.');
              placement.startSeconds        = emergencyStart;
              placement.clipDurationSeconds = duration;
              placement.resolvedAnchor      = options.slideAnchor;
              placement.allUnsafe           = true;
              unsafeFallbackCount++;
              cursor = emergencyStart + duration + slideGap;
              catPlacedCounts[catName] = (catPlacedCounts[catName] || 0) + 1;
              catPlacedIntervals[catName].push({ start: emergencyStart, end: emergencyStart + duration });
              index++;
              placeNext();
              return;
            }
            var gt = freeGaps[fgIdx++];
            previewCached(options, placement, gt, function (err, preview) {
              if (!err && preview && !preview.allUnsafe) {
                log('Slide ' + (index + 1) + ' [' + catName + ']: zone exhausted — gap placement at ' + secToMS(gt) + ' ✓');
                placement.startSeconds        = gt;
                placement.clipDurationSeconds = duration;
                placement.resolvedAnchor      = preview.resolvedAnchor || options.slideAnchor;
                placement.allUnsafe           = false;
                catPlacedCounts[catName]      = (catPlacedCounts[catName] || 0) + 1;
                catPlacedIntervals[catName].push({ start: gt, end: gt + duration });
                // Do NOT advance cursor — gap insert, not an append
                index++;
                placeNext();
                return;
              }
              tryFreeGap();
            });
          }());
          return;
        }
        // No gaps available — defer or sequential
        if (moreInCat) {
          log('Slide ' + (index + 1) + ' [' + catName + ']: zone exhausted — deferring to gap retry.');
          deferredSlides[catName].push(placement);
          index++;
          placeNext();
          return;
        }
        // Last slide in category — must place now (no more chances)
        var emergencyStart = Math.max(0, cursor);
        log('Slide ' + (index + 1) + ' [' + catName + ']: zone exhausted — appending sequentially at ' +
            secToMS(emergencyStart) + ' (zone ended ' + secToMS(hardDeadline) + ').');
        placement.startSeconds        = emergencyStart;
        placement.clipDurationSeconds = duration;
        placement.resolvedAnchor      = options.slideAnchor;
        placement.allUnsafe           = true;
        unsafeFallbackCount++;
        cursor = emergencyStart + duration + slideGap;
        catPlacedCounts[catName] = (catPlacedCounts[catName] || 0) + 1;
        catPlacedIntervals[catName].push({ start: emergencyStart, end: emergencyStart + duration });
        index++;
        placeNext();
        return;
      }

      // ── Build candidate list ─────────────────────────────────────────────
      // "target" is the closest we can be to the originally planned time
      // while still being inside [cursor, hardDeadline].
      var target = Math.max(cursor, Math.min(originalStart, hardDeadline));
      target     = nextAvailableStart(target, duration, blockedRanges);
      if (target > hardDeadline) target = hardDeadline; // edge: blocked range pushed us over

      // Interleave forward and backward steps from target so that we search
      // both directions simultaneously.  This compresses into short sequences
      // without biasing toward only forward or only backward.
      var maxSteps = Math.max(
        Math.ceil((hardDeadline - target) / skipStep),
        Math.ceil((target - cursor) / skipStep),
        1
      );
      // Cap at 80 so we can bridge large V1 gaps (e.g. 200 s / 3 s step ≈ 67 steps)
      // without iterating the entire timeline.
      maxSteps = Math.min(maxSteps, 80);

      var candidates = [];
      var seen = {};

      function addCandidate(t) {
        t = Math.round(t * 10) / 10; // 0.1s granularity
        if (t < cursor - 0.05 || t > hardDeadline + 0.05) return;
        if (overlapsBlockedRange(t, duration, blockedRanges)) return;
        var key = Math.round(t * 10);
        if (seen[key]) return;
        seen[key] = true;
        candidates.push(t);
      }

      addCandidate(target); // try the ideal position first
      for (var step = 1; step <= maxSteps; step++) {
        addCandidate(target + step * skipStep);
        addCandidate(target - step * skipStep);
      }

      // Sort by proximity to originalStart (closest first)
      candidates.sort(function (a, b) {
        return Math.abs(a - originalStart) - Math.abs(b - originalStart);
      });

      log('Slide ' + (index + 1) + ' [' + catName + ']: searching ' + candidates.length +
          ' candidates in [' + secToMS(cursor) + ' – ' + secToMS(hardDeadline) + ']...');

      // ── Sequential async probe ────────────────────────────────────────────
      var bestFallback = null;
      var attemptIdx   = 0;

      function updateBest(preview) {
        if (!preview) return;
        if (!bestFallback || combinedScore(preview) < combinedScore(bestFallback)) {
          bestFallback = preview;
        }
      }

      // After coarse 3s pass fails, a 1s fine pass scans for brief clear windows
      // (scene cuts, transitions) that the 3s grid would miss.
      var fineCandidates = null; // built lazily if coarse pass fails
      var fineIdx = 0;

      function buildFineCandidates() {
        var fine = [];
        var t = cursor;
        while (t <= hardDeadline + 0.05) {
          var key = Math.round(t * 10);
          if (!seen[key] && !overlapsBlockedRange(t, duration, blockedRanges)) {
            fine.push(Math.round(t * 10) / 10);
          }
          t += 1.0; // 1s steps
        }
        // Sort by proximity to originalStart so nearby slots are tried first
        fine.sort(function (a, b) {
          return Math.abs(a - originalStart) - Math.abs(b - originalStart);
        });
        return fine;
      }

      function tryNextCandidate() {
        // ── Coarse pass (3s steps) ────────────────────────────────────────
        if (attemptIdx < candidates.length) {
          var t = candidates[attemptIdx];
          attemptIdx++;
          log('Slide ' + (index + 1) + ': checking ' + secToMS(t) + '...');
          previewCached(options, placement, t, function (err, preview) {
            if (!err && preview) {
              updateBest(preview);
              if (!preview.allUnsafe) {
                var delta = Math.abs(t - originalStart);
                if (delta > 0.5) {
                  log('Slide ' + (index + 1) + ': ✓ safe at ' + secToMS(t) + ' (shifted ' +
                      (t > originalStart ? '+' : '-') + delta.toFixed(1) + 's). Corner: ' + preview.resolvedAnchor + '.');
                }
                finalize(preview, false);
                return;
              }
              log('Slide ' + (index + 1) + ': face/head at ' + secToMS(t) + ' → next candidate...');
            } else {
              log('Slide ' + (index + 1) + ': frame error at ' + secToMS(t) + ' → next candidate...');
            }
            tryNextCandidate();
          });
          return;
        }

        // ── Fine pass (1s steps, only if coarse pass failed) ─────────────
        // Skip fine pass entirely when the coarse best is clearly face-saturated:
        // if BOTH corners have skinRatio > 0.15, the area is fully face-covered and
        // 1-second increments will not help — stop immediately to save time.
        if (!fineCandidates) {
          var bestScores = bestFallback && bestFallback.scores;
          var bestTL = bestScores && bestScores['top-left']  ? bestScores['top-left'].skinRatio  : 1;
          var bestTR = bestScores && bestScores['top-right'] ? bestScores['top-right'].skinRatio : 1;
          if (bestFallback && bestTL > 0.15 && bestTR > 0.15) {
            log('Slide ' + (index + 1) + ': both corners face-saturated (TL:' + bestTL.toFixed(2) +
                ' TR:' + bestTR.toFixed(2) + ') — skipping fine pass.');
            finalize(bestFallback, true);
            return;
          }
          fineCandidates = buildFineCandidates().slice(0, 25); // cap at 25

          // When the coarse pass found zero candidates (entire range V1-blocked) the
          // fine pass may also be empty.  As a last resort, try the moment right after
          // each V1 range ends — those positions are guaranteed to be unblocked and
          // often correspond to natural scene changes or gaps between clips.
          if (fineCandidates.length === 0 && blockedRanges.length > 0) {
            var v1GapEdges = [];
            for (var gi = 0; gi < blockedRanges.length; gi++) {
              var gapT = Math.round(blockedRanges[gi].end * 10) / 10;
              if (gapT >= cursor - 0.05 && gapT <= hardDeadline + 0.05 &&
                  !overlapsBlockedRange(gapT, duration, blockedRanges)) {
                v1GapEdges.push(gapT);
              }
            }
            if (v1GapEdges.length > 0) {
              v1GapEdges.sort(function (a, b) { return Math.abs(a - originalStart) - Math.abs(b - originalStart); });
              fineCandidates = v1GapEdges;
              log('Slide ' + (index + 1) + ': all slots V1-blocked — trying ' + v1GapEdges.length + ' V1-gap edge positions...');
            }
          }

          if (fineCandidates.length) {
            log('Slide ' + (index + 1) + ': coarse pass done — switching to 1s fine pass (' +
                fineCandidates.length + ' positions)...');
          }
        }
        if (fineIdx < fineCandidates.length) {
          var ft = fineCandidates[fineIdx];
          fineIdx++;
          log('Slide ' + (index + 1) + ': fine-checking ' + secToMS(ft) + '...');
          previewCached(options, placement, ft, function (err, preview) {
            if (!err && preview) {
              updateBest(preview);
              if (!preview.allUnsafe) {
                var fdelta = Math.abs(ft - originalStart);
                if (fdelta > 0.5) {
                  log('Slide ' + (index + 1) + ': ✓ safe at ' + secToMS(ft) + ' (shifted ' +
                      (ft > originalStart ? '+' : '-') + fdelta.toFixed(1) + 's). Corner: ' + preview.resolvedAnchor + '.');
                }
                finalize(preview, false);
                return;
              }
              log('Slide ' + (index + 1) + ': face at ' + secToMS(ft) + ' → next fine...');
            } else {
              log('Slide ' + (index + 1) + ': frame error at ' + secToMS(ft) + ' → next fine...');
            }
            tryNextCandidate();
          });
          return;
        }

        // Both passes exhausted.  Try gap positions first before deferring or giving up.
        var moreRemain = placements.slice(index + 1).some(function (p) {
          return (p.categoryName || '') === catName;
        });
        var exhaustGaps = gapCandidates(zone.start, zone.end - duration, catPlacedIntervals[catName] || [], duration);
        if (exhaustGaps.length > 0) {
          var exIdx = 0;
          (function tryExhaustGap() {
            if (exIdx >= exhaustGaps.length) {
              // All gap positions also face-blocked — defer or use best
              if (moreRemain) {
                log('Slide ' + (index + 1) + ' [' + catName + ']: fully face-blocked, gaps unsafe — deferring.');
                deferredSlides[catName].push(placement);
                index++;
                placeNext();
                return;
              }
              log('Slide ' + (index + 1) + ': all candidates checked — using best found.');
              finalize(bestFallback, true);
              return;
            }
            var gt = exhaustGaps[exIdx++];
            previewCached(options, placement, gt, function (err, preview) {
              if (!err && preview && !preview.allUnsafe) {
                log('Slide ' + (index + 1) + ' [' + catName + ']: face-blocked — gap placement at ' + secToMS(gt) + ' ✓');
                placement.startSeconds        = gt;
                placement.clipDurationSeconds = duration;
                placement.resolvedAnchor      = preview.resolvedAnchor || options.slideAnchor;
                placement.allUnsafe           = false;
                catPlacedCounts[catName]      = (catPlacedCounts[catName] || 0) + 1;
                catPlacedIntervals[catName].push({ start: gt, end: gt + duration });
                movedCount++;
                // Do NOT advance cursor — gap insert, not an append
                index++;
                placeNext();
                return;
              }
              tryExhaustGap();
            });
          }());
          return;
        }
        if (moreRemain) {
          log('Slide ' + (index + 1) + ' [' + catName + ']: fully face-blocked — deferring to gap retry.');
          deferredSlides[catName].push(placement);
          index++;
          placeNext();
          return;
        }
        // Last slide in category — use best found (no more slides to defer past)
        log('Slide ' + (index + 1) + ': all candidates checked — using best found.');
        finalize(bestFallback, true);
      }

      function finalize(result, isFallback) {
        var finalStart  = (result && typeof result.startSeconds === 'number') ? result.startSeconds : target;
        var finalAnchor = (result && result.resolvedAnchor) ? result.resolvedAnchor : options.slideAnchor;
        var isUnsafe    = isFallback && (!result || !!result.allUnsafe);

        // ── Hard safety clamps (non-negotiable) ──────────────────────────
        if (finalStart < cursor) {
          finalStart = cursor;
          log('Slide ' + (index + 1) + ': clamped UP to cursor ' + secToMS(cursor) + '.');
        }
        if (finalStart > hardDeadline) {
          finalStart = hardDeadline;
          log('Slide ' + (index + 1) + ': clamped DOWN to deadline ' + secToMS(hardDeadline) + '.');
        }
        finalStart = Math.max(finalStart, 0);

        // ── V1 safety check — clamping can accidentally land on a V1 clip ──
        // Always nudge past a blocked range.  Placing slightly past the nominal
        // window end is far better than overlapping a V1 clip.
        if (blockedRanges.length > 0 && overlapsBlockedRange(finalStart, duration, blockedRanges)) {
          var nudged = nextAvailableStart(finalStart, duration, blockedRanges);
          log('Slide ' + (index + 1) + ': nudged out of V1 range ' + secToMS(finalStart) + ' → ' + secToMS(nudged) + '.');
          finalStart = nudged;
        }

        placement.startSeconds       = finalStart;
        placement.clipDurationSeconds = duration;
        placement.resolvedAnchor     = finalAnchor;
        placement.allUnsafe          = isUnsafe;

        if (Math.abs(finalStart - originalStart) > 0.5) movedCount++;
        if (isUnsafe) unsafeFallbackCount++;

        // Advance cursor — next slide cannot start until this one finishes
        cursor = finalStart + duration + slideGap;
        catPlacedCounts[catName]    = (catPlacedCounts[catName]    || 0) + 1;
        catPlacedIntervals[catName] = catPlacedIntervals[catName]  || [];
        catPlacedIntervals[catName].push({ start: finalStart, end: finalStart + duration });

        index++;
        placeNext();
      }

      tryNextCandidate();
    }

    placeNext();
  }

  function getGitHubReleaseApiUrl(repo) {
    return 'https://api.github.com/repos/' + repo + '/releases/latest';
  }

  function isTestUpdateChannelEnabled() {
    try {
      if (localStorage.getItem(UPDATE_CHANNEL_STORAGE_KEY) === 'test') {
        return true;
      }
    } catch (e) {}

    try {
      var root = extensionRoot || resolveExtensionRoot();
      return !!(root && fs.existsSync(path.join(root, TEST_UPDATE_FLAG_FILE)));
    } catch (e1) {
      return false;
    }
  }

  function requestLatestRelease(repo, callback) {
    requestJson(getGitHubReleaseApiUrl(repo), callback);
  }

  function requestReleases(repo, callback) {
    requestJson('https://api.github.com/repos/' + repo + '/releases?per_page=20', callback);
  }

  function getUpdateRelease(callback) {
    if (!isTestUpdateChannelEnabled()) {
      requestLatestRelease(updateRepo, callback);
      return;
    }

    requestReleases(updateRepo, function (err, releases) {
      if (err) {
        callback(err);
        return;
      }
        if (Array.isArray(releases)) {
          for (var i = 0; i < releases.length; i++) {
            var release = releases[i];
            if (!release || release.draft || !release.prerelease) continue;
            if (getReleaseZipAsset(release) && compareVersions(getReleaseVersion(release), updateState.installedVersion) >= 0) {
              callback(null, release);
              return;
            }
          }
        }
      requestLatestRelease(updateRepo, callback);
    });
  }

  function getTempPath(name) {
    return path.join(os.tmpdir(), 'smtv-slides-updater', String(name || ''));
  }

  function stripWindowsExtendedPathPrefix(filePath) {
    return process.platform === 'win32'
      ? String(filePath || '').replace(/^\\\\\?\\/, '')
      : String(filePath || '');
  }

  function quotePowerShellLiteral(str) {
    return "'" + String(str || '').replace(/'/g, "''") + "'";
  }

  function ensureDir(dirPath) {
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath, { recursive: true });
    }
  }

  function removeDirRecursive(dirPath) {
    if (!fs.existsSync(dirPath)) return;
    if (fs.rmSync) {
      fs.rmSync(dirPath, { recursive: true, force: true });
      return;
    }
    fs.readdirSync(dirPath).forEach(function (entry) {
      var fullPath = path.join(dirPath, entry);
      var stat = fs.lstatSync(fullPath);
      if (stat.isDirectory()) {
        removeDirRecursive(fullPath);
      } else {
        fs.unlinkSync(fullPath);
      }
    });
    fs.rmdirSync(dirPath);
  }

  function removeFileOrDir(targetPath) {
    if (!fs.existsSync(targetPath)) return;
    var stat = fs.lstatSync(targetPath);
    if (stat.isDirectory()) {
      removeDirRecursive(targetPath);
    } else {
      fs.unlinkSync(targetPath);
    }
  }

  function copyDirRecursive(srcDir, destDir) {
    ensureDir(destDir);
    fs.readdirSync(srcDir).forEach(function (entry) {
      var srcPath = path.join(srcDir, entry);
      var destPath = path.join(destDir, entry);
      var stat = fs.lstatSync(srcPath);
      if (stat.isDirectory()) {
        copyDirRecursive(srcPath, destPath);
      } else {
        fs.copyFileSync(srcPath, destPath);
      }
    });
  }

  function clearDirectoryContents(dirPath) {
    if (!fs.existsSync(dirPath)) return;
    fs.readdirSync(dirPath).forEach(function (entry) {
      removeFileOrDir(path.join(dirPath, entry));
    });
  }

  function requestJson(url, callback, redirectCount) {
    var redirects = redirectCount || 0;
    https.get(url, {
      headers: {
        'User-Agent': 'SMTV-Slides-Updater',
        'Accept': 'application/vnd.github+json'
      }
    }, function (res) {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location && redirects < 5) {
        res.resume();
        requestJson(res.headers.location, callback, redirects + 1);
        return;
      }
      if (res.statusCode !== 200) {
        res.resume();
        callback(new Error('GitHub update check failed with status ' + res.statusCode + '.'));
        return;
      }
      var body = '';
      res.setEncoding('utf8');
      res.on('data', function (chunk) { body += chunk; });
      res.on('end', function () {
        try {
          callback(null, JSON.parse(body));
        } catch (e) {
          callback(new Error('Could not parse GitHub release response.'));
        }
      });
    }).on('error', function (err) {
      callback(err);
    });
  }

  function downloadFile(url, destPath, callback, redirectCount) {
    var redirects = redirectCount || 0;
    ensureDir(path.dirname(destPath));
    https.get(url, {
      headers: {
        'User-Agent': 'SMTV-Slides-Updater',
        'Accept': 'application/octet-stream'
      }
    }, function (res) {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location && redirects < 5) {
        res.resume();
        downloadFile(res.headers.location, destPath, callback, redirects + 1);
        return;
      }
      if (res.statusCode !== 200) {
        res.resume();
        callback(new Error('Download failed with status ' + res.statusCode + '.'));
        return;
      }

      var file = fs.createWriteStream(destPath);
      res.pipe(file);
      file.on('finish', function () {
        file.close(function () { callback(null); });
      });
      file.on('error', function (err) {
        try { file.close(function () {}); } catch (e) {}
        callback(err);
      });
    }).on('error', function (err) {
      callback(err);
    });
  }

  function extractZip(zipPath, destDir) {
    removeDirRecursive(destDir);
    ensureDir(destDir);
    if (process.platform === 'win32') {
      childProcess.execFileSync('powershell.exe', ['-NoProfile', '-Command', 'Expand-Archive -LiteralPath "' + zipPath.replace(/"/g, '""') + '" -DestinationPath "' + destDir.replace(/"/g, '""') + '" -Force']);
      return;
    }
    childProcess.execFileSync('unzip', ['-oq', zipPath, '-d', destDir]);
  }

  function findExtensionRoot(dirPath, depth) {
    var maxDepth = typeof depth === 'number' ? depth : 4;
    if (!fs.existsSync(dirPath) || maxDepth < 0) return '';
    if (fs.existsSync(path.join(dirPath, 'CSXS', 'manifest.xml'))) {
      return dirPath;
    }
    var entries = fs.readdirSync(dirPath, { withFileTypes: true });
    for (var i = 0; i < entries.length; i++) {
      if (!entries[i].isDirectory()) continue;
      var nested = findExtensionRoot(path.join(dirPath, entries[i].name), maxDepth - 1);
      if (nested) return nested;
    }
    return '';
  }

  function validateExtractedExtension(extractedRoot) {
    var manifestFile = path.join(extractedRoot, 'CSXS', 'manifest.xml');
    if (!fs.existsSync(manifestFile)) {
      throw new Error('Downloaded update does not contain CSXS/manifest.xml.');
    }
    var expectedBundleId = readManifestBundleId(manifestPath);
    var actualBundleId = readManifestBundleId(manifestFile);
    if (expectedBundleId && actualBundleId && expectedBundleId !== actualBundleId) {
      throw new Error('Downloaded update is for a different extension bundle.');
    }
  }

  function installExtractedExtension(extractedRoot) {
    validateExtractedExtension(extractedRoot);
    clearDirectoryContents(extensionRoot);
    copyDirRecursive(extractedRoot, extensionRoot);
  }

  function launchWindowsDeferredInstaller(extractedRoot, tempRoot, latestVersion) {
    validateExtractedExtension(extractedRoot);

    var installerScriptPath = path.join(tempRoot, 'install-update.ps1');
    var scriptLines = [
      "$ErrorActionPreference = 'Stop'",
      '$SourceDir = ' + quotePowerShellLiteral(stripWindowsExtendedPathPrefix(extractedRoot)),
      '$TargetDir = ' + quotePowerShellLiteral(stripWindowsExtendedPathPrefix(extensionRoot)),
      '$TempRoot = ' + quotePowerShellLiteral(stripWindowsExtendedPathPrefix(tempRoot)),
      '$StatusFile = ' + quotePowerShellLiteral(stripWindowsExtendedPathPrefix(updateInstallStatusFile)),
      '$Version = ' + quotePowerShellLiteral(latestVersion || ''),
      '',
      'function Write-UpdateStatus($State, $Message) {',
      '  @{',
      "    state = $State",
      "    version = $Version",
      "    message = $Message",
      "    updatedAt = (Get-Date).ToString('o')",
      '  } | ConvertTo-Json | Set-Content -LiteralPath $StatusFile -Encoding UTF8',
      '}',
      '',
      'function Wait-ForPremiereExit {',
      '  $deadline = (Get-Date).AddMinutes(30)',
      '  while ((Get-Date) -lt $deadline) {',
      "    $running = Get-Process -Name 'Adobe Premiere Pro' -ErrorAction SilentlyContinue",
      '    if (-not $running) { return }',
      '    Start-Sleep -Seconds 2',
      '  }',
      "  throw 'Premiere Pro did not close in time for the update to finish.'",
      '}',
      '',
      'function Install-UpdateFiles {',
      '  if (-not (Test-Path -LiteralPath $TargetDir)) {',
      '    New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null',
      '  }',
      '',
      '  for ($i = 0; $i -lt 180; $i++) {',
      '    try {',
      '      Get-ChildItem -LiteralPath $TargetDir -Force -ErrorAction SilentlyContinue | ForEach-Object {',
      '        Remove-Item -LiteralPath $_.FullName -Recurse -Force',
      '      }',
      '      Get-ChildItem -LiteralPath $SourceDir -Force | ForEach-Object {',
      '        Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $TargetDir $_.Name) -Recurse -Force',
      '      }',
      '      return',
      '    } catch {',
      '      Start-Sleep -Seconds 1',
      '    }',
      '  }',
      "  throw 'Could not replace the extension files after waiting for Premiere Pro to close.'",
      '}',
      '',
      'try {',
      "  Write-UpdateStatus 'pending' ('Waiting for Premiere Pro to close before installing version ' + $Version + '.')",
      '  Wait-ForPremiereExit',
      '  Install-UpdateFiles',
      "  Write-UpdateStatus 'success' ('Version ' + $Version + ' was installed successfully.')",
      '  try { Remove-Item -LiteralPath $TempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}',
      '} catch {',
      "  Write-UpdateStatus 'failed' $_.Exception.Message",
      "  Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue | Out-Null",
      "  [System.Windows.MessageBox]::Show('SMTV Auto Slides update failed: ' + $_.Exception.Message, 'SMTV Auto Slides Updater') | Out-Null",
      '  exit 1',
      '}'
    ];

    fs.writeFileSync(installerScriptPath, scriptLines.join('\r\n'), 'utf8');

    var launchCommand = [
      'Start-Process',
      '-FilePath', quotePowerShellLiteral('powershell.exe'),
      '-Verb', 'RunAs',
      '-WindowStyle', 'Hidden',
      '-ArgumentList',
      '@(' +
        quotePowerShellLiteral('-NoProfile') + ',' +
        quotePowerShellLiteral('-ExecutionPolicy') + ',' +
        quotePowerShellLiteral('Bypass') + ',' +
        quotePowerShellLiteral('-File') + ',' +
        quotePowerShellLiteral(stripWindowsExtendedPathPrefix(installerScriptPath)) +
      ')'
    ].join(' ');

    childProcess.execFileSync('powershell.exe', ['-NoProfile', '-Command', launchCommand]);
  }

  function getReleaseVersion(release) {
    return normalizeVersion((release && (release.tag_name || release.name)) || '');
  }

  function getReleaseZipAsset(release) {
    var assets = (release && release.assets) || [];
    for (var i = 0; i < assets.length; i++) {
      var asset = assets[i];
      if (asset && asset.browser_download_url && /\.zip$/i.test(asset.name || asset.browser_download_url)) {
        return asset;
      }
    }
    return null;
  }

  function checkForUpdates(options) {
    var opts = options || {};
    if (!updateRepo) {
      updateState.latestRelease = null;
      updateState.latestVersion = '';
      setUpdateStatus('No GitHub update source is configured.');
      setUpdateUiState();
      return;
    }

    updateState.checking = true;
    setUpdateStatus(opts.silent ? 'Checking for updates in background...' : 'Checking for updates...');
    setUpdateUiState();

    getUpdateRelease(function (err, release) {
      updateState.checking = false;
      if (err) {
        updateState.latestRelease = null;
        updateState.latestVersion = '';
        setUpdateStatus('Update check failed: ' + err.message);
        setUpdateUiState();
        return;
      }

        var latestVersion = getReleaseVersion(release);
        var zipAsset = getReleaseZipAsset(release);
        var allowTestReinstall = isLocalPrereleaseReinstallAvailable(release);
        updateState.latestVersion = latestVersion || '';
        updateState.latestRelease = zipAsset ? release : null;
        persistUpdateInfo(latestVersion);

        if (!zipAsset) {
          setUpdateStatus('A release was found, but no zip asset is available to install.');
        } else if (compareVersions(latestVersion, updateState.installedVersion) > 0) {
          setUpdateStatus('Version ' + latestVersion + ' is available. Click Update Now to install it.');
        } else if (allowTestReinstall) {
          setUpdateStatus('Test prerelease ' + latestVersion + ' is available for reinstall on this machine. Click Update Now to test the updater.');
        } else {
          setUpdateStatus('You are up to date.');
        }
      setUpdateUiState();
    });
  }

  function installLatestUpdate() {
    if (!updateRepo) {
      setUpdateStatus('No GitHub update source is configured.');
      return;
    }

    var release = updateState.latestRelease;
    var latestVersion = updateState.latestVersion;
    var zipAsset = getReleaseZipAsset(release);
    if (!release || !zipAsset) {
      setUpdateStatus('No downloadable update is ready yet. Restart the extension or wait for the startup check to finish.');
      return;
    }

      showUpdateModal(
        'Install Update',
        buildUpdateNotesMessage(release, 'Install version ' + latestVersion + ' from ' + updateRepo + ' now?\nPremiere Pro should be restarted after the update.', { popupSummary: true }),
        { confirm: true, okText: 'Install', cancelText: 'Cancel' }
      ).then(function (confirmed) {
      if (!confirmed) {
        return;
      }

      updateState.installing = true;
      setUpdateStatus('Downloading update...');
      setUpdateUiState();
      savePendingUpdateInfo(latestVersion, release.name || release.tag_name || latestVersion, getReleaseNotes(release));

      var tempRoot = getTempPath(String(Date.now()));
      var zipPath = path.join(tempRoot, 'update.zip');
      var extractPath = path.join(tempRoot, 'extracted');

      try {
        ensureDir(tempRoot);
      } catch (e) {
        updateState.installing = false;
        setUpdateStatus('Could not prepare temp update folder: ' + e.message);
        setUpdateUiState();
        return;
      }

      downloadFile(zipAsset.browser_download_url, zipPath, function (downloadErr) {
        if (downloadErr) {
          updateState.installing = false;
          setUpdateStatus('Update download failed: ' + downloadErr.message);
          setUpdateUiState();
          return;
        }

        try {
          setUpdateStatus('Extracting update...');
          extractZip(zipPath, extractPath);
          var extractedExtensionRoot = findExtensionRoot(extractPath);
          if (!extractedExtensionRoot) {
            throw new Error('Could not find the extension root in the downloaded zip.');
          }

          if (process.platform === 'win32') {
            setUpdateStatus('Preparing update installer...');
            saveUpdateInstallStatus({
              state: 'staged',
              version: latestVersion,
              message: 'Update is staged and waiting for Windows approval.',
              updatedAt: new Date().toISOString()
            });
            launchWindowsDeferredInstaller(extractedExtensionRoot, tempRoot, latestVersion);
            updateState.installing = false;
            setUpdateStatus('Update is staged. Accept the Windows prompt, then close Premiere Pro and wait a few seconds before reopening. The installer will finish after Premiere exits and update to version ' + latestVersion + '.');
            setUpdateUiState();
            return;
          }

          setUpdateStatus('Installing update...');
          installExtractedExtension(extractedExtensionRoot);
          updateState.installedVersion = readManifestVersion(manifestPath) || latestVersion;
          updateState.installing = false;
          clearUpdateInstallStatus();
          updateState.latestRelease = compareVersions(updateState.latestVersion, updateState.installedVersion) > 0 ? updateState.latestRelease : null;
          setUpdateStatus('Update installed. Please restart Premiere Pro to load version ' + updateState.installedVersion + '.');
          setUpdateUiState();
          showUpdateModal('Update Installed', buildUpdateNotesMessage(release, 'The update was installed successfully.', { popupSummary: true }), { okText: 'OK' }).then(function () {
            clearPendingUpdateInfo();
          });
        } catch (installErr) {
          updateState.installing = false;
          setUpdateStatus('Update install failed: ' + installErr.message);
          setUpdateUiState();
        }
      });
    });
  }

  function escapeForEval(str) {
    return String(str)
      .replace(/\\/g, '\\\\')
      .replace(/'/g, "\\'")
      .replace(/\r/g, '')
      .replace(/\n/g, '\\n');
  }

  function normalizeToken(str) {
    return String(str || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  }

  function normalizeTitleKey(str) {
    return String(str || '').toLowerCase().replace(/\.mov$/i, '').replace(/[^a-z0-9]+/g, ' ').trim();
  }

  function formatDisplayTitle(str) {
    return String(str || '')
      .replace(/[_\s]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  var canonicalLanguageAliasMap = {
    English: ['english', 'eng'],
    Arabic: ['arabic', 'ara'],
    Aulacese: ['aulacese', 'aulac', 'au', 'aul', 'vietnamese', 'viet', 'vie'],
    Bulgarian: ['bulgarian', 'bul'],
    Chinese: ['chinese', 'chi', 'zho', 'zh'],
    'Chinese Simplified': ['chinese simplified', 'chi simp', 'chisimp', 'simplified chinese'],
    'Chinese Traditional': ['chinese traditional', 'chi trad', 'chitrad', 'traditional chinese'],
    Croatian: ['croatian', 'cro', 'hrv'],
    Czech: ['czech', 'cze', 'ces'],
    Dutch: ['dutch', 'nederlands', 'ned', 'nld'],
    Estonian: ['estonian', 'est'],
    Ewe: ['ewe'],
    Finnish: ['finnish', 'fin'],
    French: ['french', 'fre', 'fra'],
    German: ['german', 'ger', 'deu'],
    Greek: ['greek', 'gre', 'ell'],
    Hebrew: ['hebrew', 'heb'],
    Hindi: ['hindi', 'hin'],
    Hungarian: ['hungarian', 'hun'],
    Indonesian: ['indonesian', 'ind', 'ina'],
    Italian: ['italian', 'ita'],
    Japanese: ['japanese', 'jap', 'jpn'],
    Korean: ['korean', 'kor'],
    Malay: ['malay', 'malaysian', 'mal'],
    Mongolian: ['mongolian', 'mon'],
    Norwegian: ['norwegian', 'norway', 'nor'],
    Persian: ['persian', 'per', 'fas'],
    Polish: ['polish', 'pol'],
    Portuguese: ['portuguese', 'por'],
    Punjabi: ['punjabi', 'pun', 'pan'],
    Romanian: ['romanian', 'rom', 'ron'],
    Russian: ['russian', 'rus'],
    Serbian: ['serbian', 'srp', 'scc'],
    Slovenian: ['slovenian', 'slv'],
    Spanish: ['spanish', 'spa'],
    Swedish: ['swedish', 'swe'],
    Telugu: ['telugu', 'telegu', 'tel'],
    Thai: ['thai', 'tha'],
    Turkish: ['turkish', 'tur'],
    Ukrainian: ['ukrainian', 'ukr'],
    Urdu: ['urdu', 'urd'],
    Zulu: ['zulu', 'zul']
  };
  var canonicalLanguageAliasLookup = null;

  function titleCaseWords(str) {
    return String(str || '').replace(/\b[a-z]/g, function (ch) { return ch.toUpperCase(); });
  }

  function getCanonicalLanguageAliasLookup() {
    if (canonicalLanguageAliasLookup) return canonicalLanguageAliasLookup;

    canonicalLanguageAliasLookup = {};
    for (var canonicalName in canonicalLanguageAliasMap) {
      if (!canonicalLanguageAliasMap.hasOwnProperty(canonicalName)) continue;
      var aliases = canonicalLanguageAliasMap[canonicalName].slice();
      aliases.push(normalizeToken(canonicalName));
      for (var i = 0; i < aliases.length; i++) {
        var alias = normalizeToken(aliases[i]);
        if (!alias) continue;
        canonicalLanguageAliasLookup[alias] = canonicalName;
        canonicalLanguageAliasLookup[alias.replace(/\s+/g, '')] = canonicalName;
      }
    }

    return canonicalLanguageAliasLookup;
  }

  function canonicalizeLanguageName(languageName) {
    var langNorm = normalizeToken(languageName);
    var compactNorm = langNorm.replace(/\s+/g, '');
    if (!langNorm) return null;

    var aliasLookup = getCanonicalLanguageAliasLookup();
    if (aliasLookup[langNorm]) return aliasLookup[langNorm];
    if (aliasLookup[compactNorm]) return aliasLookup[compactNorm];

    var partialMatches = [];
    for (var alias in aliasLookup) {
      if (!aliasLookup.hasOwnProperty(alias) || !alias) continue;
      if (compactNorm === alias) {
        return aliasLookup[alias];
      }
      if (compactNorm.indexOf(alias) !== -1 || alias.indexOf(compactNorm) !== -1) {
        if (partialMatches.indexOf(aliasLookup[alias]) === -1) {
          partialMatches.push(aliasLookup[alias]);
        }
      }
    }
    if (partialMatches.length === 1) {
      return partialMatches[0];
    }

    return titleCaseWords(langNorm);
  }

  function getLanguageAliases(languageFolderName) {
    var canonicalName = canonicalizeLanguageName(languageFolderName);
    var aliases = canonicalLanguageAliasMap[canonicalName];
    if (aliases && aliases.length) return aliases.slice();

    var langNorm = normalizeToken(languageFolderName);
    return langNorm ? [langNorm] : [];
  }

  function getCanonicalSaveTheEarthTitleKey(title) {
    var key = normalizeTitleKey(title);
    var stripped = key.replace(/\s*\d+$/, '').trim();
    var compact = stripped.replace(/\s+/g, '');

    if (compact === 'bekind') return 'be kind';
    if (compact === 'befrugal') return 'be frugal';

    var greenTokens = stripped.split(' ').filter(function (token) { return !!token; }).sort().join(' ');
    if (greenTokens === 'be go green veg') return 'be veg go green';

    return stripped || key;
  }

  function getCanonicalGroupingTitleKey(title, categoryName) {
    if (categoryName === 'Be Vegan Keep Peace') {
      return 'be vegan keep peace';
    }

    var key = normalizeTitleKey(title);
    if (!key) return '';

    key = key.replace(/\b(the|a|an)\b/g, ' ');
    key = key.replace(/\s+/g, ' ').trim();

    if (categoryName === 'Save the Earth') {
      return getCanonicalSaveTheEarthTitleKey(key);
    }

    return key;
  }

  function choosePreferredDisplayTitle(existingTitle, candidateTitle) {
    var existing = formatDisplayTitle(existingTitle);
    var candidate = formatDisplayTitle(candidateTitle);
    if (!existing) return candidate;
    if (!candidate) return existing;

    var existingNorm = normalizeTitleKey(existing);
    var candidateNorm = normalizeTitleKey(candidate);
    var existingHasLeadingArticle = /^(the|a|an)\b/.test(existingNorm);
    var candidateHasLeadingArticle = /^(the|a|an)\b/.test(candidateNorm);

    if (existingHasLeadingArticle !== candidateHasLeadingArticle) {
      return candidateHasLeadingArticle ? existing : candidate;
    }

    if (candidate.length < existing.length) return candidate;
    return existing;
  }

  function getSaveTheEarthFamilyKey(title, titlesMap) {
    var key = normalizeTitleKey(title);
    var canonical = getCanonicalSaveTheEarthTitleKey(title);
    if (!canonical || canonical === key) return canonical || key;

    var titles = Object.keys(titlesMap || {});
    var hasRelatedVariant = titles.some(function (candidate) {
      var candidateKey = normalizeTitleKey(candidate);
      if (candidateKey === key) return false;
      return getCanonicalSaveTheEarthTitleKey(candidate) === canonical;
    });

    return hasRelatedVariant ? canonical : key;
  }

  function getSaveTheEarthFallbackTitles(title, titlesMap) {
    var familyKey = getSaveTheEarthFamilyKey(title, titlesMap);
    var familyTitles = Object.keys(titlesMap || {}).filter(function (candidate) {
      return getSaveTheEarthFamilyKey(candidate, titlesMap) === familyKey;
    });
    return familyTitles.length ? familyTitles : [title];
  }

  function parseTitleFromFile(filePath, languageFolderName) {
    var ext = path.extname(filePath).toLowerCase();
    if (ext !== '.mov') return null;

    var base = path.basename(filePath, ext);
    var underscoreIndex = base.indexOf('_');
    if (underscoreIndex === -1 || underscoreIndex === base.length - 1) return null;

    var prefix = normalizeToken(base.substring(0, underscoreIndex));
    var title = base.substring(underscoreIndex + 1).trim();
    if (!title) return null;

    var allowedTokens = getLanguageAliases(languageFolderName);

    var matched = false;
    for (var i = 0; i < allowedTokens.length; i++) {
      if (allowedTokens[i] && prefix.indexOf(allowedTokens[i]) !== -1) {
        matched = true;
        break;
      }
    }
    if (!matched) return null;

    return title;
  }

  function parseFlatBeVeganLanguage(fileName) {
    var ext = path.extname(fileName).toLowerCase();
    if (ext !== '.mov') return null;
    var base = path.basename(fileName, ext);
    var parts = base.split('_');
    if (parts.length < 3) return null;
    var langCode = parts[parts.length - 1].trim();
    if (!langCode) return null;
    return canonicalizeLanguageName(langCode);
  }

  function createEmptyScanResult(categoryName) {
    return {
      categoryName: categoryName,
      isFlatSingleTitle: categoryName === 'Be Vegan Keep Peace',
      languageDirs: [],
      titlesMap: {},
      titleDisplayMap: {}
    };
  }

  function addEntryToScanResult(scanResult, title, languageName, filePath, categoryName) {
    var canonicalLanguage = canonicalizeLanguageName(languageName);
    var titleKey = getCanonicalGroupingTitleKey(title, categoryName);
    if (!titleKey || !canonicalLanguage || !filePath) return;

    if (!scanResult.titlesMap[titleKey]) scanResult.titlesMap[titleKey] = {};
    if (!scanResult.titlesMap[titleKey][canonicalLanguage]) {
      scanResult.titlesMap[titleKey][canonicalLanguage] = filePath;
    }

    scanResult.titleDisplayMap[titleKey] = choosePreferredDisplayTitle(scanResult.titleDisplayMap[titleKey], title);

    if (scanResult.languageDirs.indexOf(canonicalLanguage) === -1) {
      scanResult.languageDirs.push(canonicalLanguage);
    }
  }

  function mergeScanResults(baseScanResult, additionScanResult) {
    var merged = baseScanResult || { categoryName: '', isFlatSingleTitle: false, languageDirs: [], titlesMap: {}, titleDisplayMap: {} };
    if (!additionScanResult) return merged;

    if (additionScanResult.isFlatSingleTitle) {
      merged.isFlatSingleTitle = true;
    }

    additionScanResult.languageDirs.forEach(function (lang) {
      if (merged.languageDirs.indexOf(lang) === -1) {
        merged.languageDirs.push(lang);
      }
    });

    Object.keys(additionScanResult.titlesMap || {}).forEach(function (title) {
      var languageMap = additionScanResult.titlesMap[title] || {};
      Object.keys(languageMap).forEach(function (lang) {
        addEntryToScanResult(
          merged,
          (additionScanResult.titleDisplayMap && additionScanResult.titleDisplayMap[title]) || title,
          lang,
          languageMap[lang],
          merged.categoryName || additionScanResult.categoryName || ''
        );
      });
    });

    return merged;
  }

  function parseFlexibleCategoryFile(filePath) {
    var ext = path.extname(filePath).toLowerCase();
    if (ext !== '.mov') return null;

    var base = path.basename(filePath, ext).trim();
    if (!base) return null;

    var match = base.match(/^Be Vegan[_\s]+Keep Peace_(.+)$/i);
    if (match) {
      return {
        categoryName: 'Be Vegan Keep Peace',
        title: 'Be Vegan Keep Peace',
        language: match[1].trim()
      };
    }

    match = base.match(/^slides\s+peace\s+2019\s+(.+?)_(.+)$/i);
    if (match) {
      return {
        categoryName: 'NEW PEACE MAKER',
        title: match[2].trim(),
        language: match[1].trim()
      };
    }

    match = base.match(/^slides\s+forgiveness\s+(.+?)_(.+)$/i);
    if (match) {
      return {
        categoryName: 'Forgiveness',
        title: match[2].trim(),
        language: match[1].trim()
      };
    }

    match = base.match(/^slides\s+save(?:the|our)e?arth\s+(.+?)_(.+)$/i);
    if (match) {
      return {
        categoryName: 'Save the Earth',
        title: match[2].trim(),
        language: match[1].trim()
      };
    }

    match = base.match(/^slides\s+veg\s+(.+?)_(.+)$/i);
    if (match) {
      return {
        categoryName: 'Veganism',
        title: match[2].trim(),
        language: match[1].trim()
      };
    }

    return null;
  }

  function scanFlexibleCategories(rootFolder) {
    var scanByCategory = {};
    var pendingDirs = [rootFolder];

    while (pendingDirs.length) {
      var currentDir = pendingDirs.pop();
      var dirents = fs.readdirSync(currentDir, { withFileTypes: true });

      dirents.forEach(function (dirent) {
        var fullPath = path.join(currentDir, dirent.name);
        if (dirent.isDirectory()) {
          if (!ignoredFolderNames[dirent.name]) {
            pendingDirs.push(fullPath);
          }
          return;
        }

        if (!dirent.isFile() || path.extname(dirent.name).toLowerCase() !== '.mov') return;

        var parsed = parseFlexibleCategoryFile(fullPath);
        if (!parsed || categoryOrder.indexOf(parsed.categoryName) === -1) return;

        if (!scanByCategory[parsed.categoryName]) {
          scanByCategory[parsed.categoryName] = createEmptyScanResult(parsed.categoryName);
        }

        addEntryToScanResult(scanByCategory[parsed.categoryName], parsed.title, parsed.language, fullPath, parsed.categoryName);
      });
    }

    return scanByCategory;
  }

  function scanSingleCategory(categoryName, categoryRoot) {
    var scanResult = createEmptyScanResult(categoryName);

    if (categoryName === 'Be Vegan Keep Peace') {
      var flatFiles = fs.readdirSync(categoryRoot, { withFileTypes: true })
        .filter(function (d) { return d.isFile() && path.extname(d.name).toLowerCase() === '.mov'; })
        .map(function (d) { return path.join(categoryRoot, d.name); });

      flatFiles.forEach(function (fullPath) {
        var lang = parseFlatBeVeganLanguage(fullPath);
        if (!lang) return;
        addEntryToScanResult(scanResult, 'Be Vegan Keep Peace', lang, fullPath, categoryName);
      });

      return scanResult;
    }

    var dirents = fs.readdirSync(categoryRoot, { withFileTypes: true });
    var languageDirs = dirents
      .filter(function (d) { return d.isDirectory() && !ignoredFolderNames[d.name]; })
      .map(function (d) {
        return {
          name: d.name,
          canonicalName: canonicalizeLanguageName(d.name)
        };
      });

    languageDirs.forEach(function (langInfo) {
      var langDir = path.join(categoryRoot, langInfo.name);
      var files = fs.readdirSync(langDir, { withFileTypes: true })
        .filter(function (d) { return d.isFile() && path.extname(d.name).toLowerCase() === '.mov'; })
        .map(function (d) { return path.join(langDir, d.name); });

      files.forEach(function (fullPath) {
        var title = parseTitleFromFile(fullPath, langInfo.name);
        if (!title) return;
        addEntryToScanResult(scanResult, title, langInfo.canonicalName, fullPath, categoryName);
      });
    });
    return scanResult;
  }

  function scanAllCategories(rootFolder) {
    var foundByCategory = {};

    categoryOrder.forEach(function (categoryName) {
      var categoryPath = path.join(rootFolder, categoryName);
      if (fs.existsSync(categoryPath) && fs.statSync(categoryPath).isDirectory()) {
        foundByCategory[categoryName] = {
          name: categoryName,
          path: categoryPath,
          scanResult: scanSingleCategory(categoryName, categoryPath)
        };
      }
    });

    var flexibleScanByCategory = scanFlexibleCategories(rootFolder);
    Object.keys(flexibleScanByCategory).forEach(function (categoryName) {
      if (!foundByCategory[categoryName]) {
        foundByCategory[categoryName] = {
          name: categoryName,
          path: rootFolder,
          scanResult: createEmptyScanResult(categoryName)
        };
      }

      foundByCategory[categoryName].scanResult = mergeScanResults(
        foundByCategory[categoryName].scanResult,
        flexibleScanByCategory[categoryName]
      );
    });

    return categoryOrder
      .filter(function (categoryName) { return !!foundByCategory[categoryName]; })
      .map(function (categoryName) { return foundByCategory[categoryName]; })
      .filter(function (categoryData) {
        return Object.keys(categoryData.scanResult.titlesMap || {}).length > 0;
      });
  }

  function chooseBatchForCategory(categoryName, scanResult, requestedCount, tracking) {
    tracking.categories = tracking.categories || {};
    tracking.categories[categoryName] = tracking.categories[categoryName] || { usedEnglishTitlesCycle: [], isFlatSingleTitle: !!scanResult.isFlatSingleTitle };
    tracking.usedLanguagesGlobalCycle = Array.isArray(tracking.usedLanguagesGlobalCycle)
      ? tracking.usedLanguagesGlobalCycle
      : [];
    tracking._currentRunUsedLanguages = Array.isArray(tracking._currentRunUsedLanguages)
      ? tracking._currentRunUsedLanguages
      : [];

    var categoryTracking = tracking.categories[categoryName];
    categoryTracking.usedEnglishTitlesCycle = Array.isArray(categoryTracking.usedEnglishTitlesCycle)
      ? categoryTracking.usedEnglishTitlesCycle
      : [];

    categoryTracking.isFlatSingleTitle = !!scanResult.isFlatSingleTitle;

    var titlesMap = scanResult.titlesMap;
    var allEnglishTitles = Object.keys(titlesMap).filter(function (title) {
      return !!titlesMap[title].English;
    });

    if (!allEnglishTitles.length) {
      throw new Error('Category "' + categoryName + '" has no usable English files.');
    }

    var unusedEnglishTitles = allEnglishTitles.filter(function (title) {
      return categoryTracking.usedEnglishTitlesCycle.indexOf(title) === -1;
    });

    if (!unusedEnglishTitles.length) {
      if (scanResult.isFlatSingleTitle) {
        unusedEnglishTitles = allEnglishTitles.slice();
      } else {
        categoryTracking.usedEnglishTitlesCycle = [];
        unusedEnglishTitles = allEnglishTitles.slice();
        log(categoryName + ': all English titles were already used once. Starting a fresh title cycle for this category.');
      }
    }

    var otherNeeded = Math.max(0, requestedCount - 1);

    function getLanguageToFileForTitle(title) {
      var baseMap = {};
      var sourceMap = titlesMap[title] || {};
      Object.keys(sourceMap).forEach(function (lang) { baseMap[lang] = sourceMap[lang]; });

      if (categoryName === 'Save the Earth') {
        var familyTitles = getSaveTheEarthFallbackTitles(title, titlesMap);
        var currentOtherCount = Object.keys(baseMap).filter(function (lang) { return lang !== 'English'; }).length;
        if (currentOtherCount < otherNeeded) {
          familyTitles.forEach(function (altTitle) {
            if (altTitle === title || !titlesMap[altTitle]) return;
            var altMap = titlesMap[altTitle];
            Object.keys(altMap).forEach(function (lang) {
              if (lang === 'English') return;
              if (!baseMap[lang]) {
                baseMap[lang] = altMap[lang];
              }
            });
          });
        }
      }

      return baseMap;
    }

    function getOtherLanguagesForTitle(title) {
      return Object.keys(getLanguageToFileForTitle(title)).filter(function (lang) { return lang !== 'English'; });
    }

    function getUnusedLanguagesForTitle(title) {
      return getOtherLanguagesForTitle(title).filter(function (lang) {
        return tracking.usedLanguagesGlobalCycle.indexOf(lang) === -1;
      });
    }

    function getUnusedThisRunLanguagesForTitle(title) {
      return getOtherLanguagesForTitle(title).filter(function (lang) {
        return tracking._currentRunUsedLanguages.indexOf(lang) === -1;
      });
    }

    function getSaveTheEarthReuseMap(title, selectedLanguages) {
      var reuseMap = {};
      if (categoryName !== 'Save the Earth') return reuseMap;

      var selectedLookup = {};
      selectedLanguages.forEach(function (lang) { selectedLookup[lang] = true; });
      tracking._currentRunUsedLanguages.forEach(function (lang) { selectedLookup[lang] = true; });

      tracking.usedLanguagesGlobalCycle.forEach(function (lang) {
        if (lang === 'English' || selectedLookup[lang]) return;

        var matchingTitle = Object.keys(titlesMap).find(function (candidateTitle) {
          return !!(titlesMap[candidateTitle] && titlesMap[candidateTitle][lang]);
        });

        if (matchingTitle) {
          reuseMap[lang] = titlesMap[matchingTitle][lang];
          selectedLookup[lang] = true;
        }
      });

      return reuseMap;
    }

    var candidatePool = shuffle(unusedEnglishTitles).sort(function (a, b) {
      var bUnusedThisRun = getUnusedThisRunLanguagesForTitle(b).length;
      var aUnusedThisRun = getUnusedThisRunLanguagesForTitle(a).length;
      if (bUnusedThisRun !== aUnusedThisRun) return bUnusedThisRun - aUnusedThisRun;

      var bUnused = getUnusedLanguagesForTitle(b).length;
      var aUnused = getUnusedLanguagesForTitle(a).length;
      if (bUnused !== aUnused) return bUnused - aUnused;

      return getOtherLanguagesForTitle(b).length - getOtherLanguagesForTitle(a).length;
    });

    var chosenTitle = candidatePool[0];
    var chosenTitleDisplay = (scanResult.titleDisplayMap && scanResult.titleDisplayMap[chosenTitle]) || chosenTitle;
    var languageToFile = getLanguageToFileForTitle(chosenTitle);
    var otherLanguages = Object.keys(languageToFile).filter(function (lang) { return lang !== 'English'; });
    var pickedOtherLanguages = shuffle(otherLanguages.filter(function (lang) {
      return tracking._currentRunUsedLanguages.indexOf(lang) === -1 &&
        tracking.usedLanguagesGlobalCycle.indexOf(lang) === -1;
    })).slice(0, otherNeeded);
    var warningMessages = [];
    var reusedFreshness = false;
    var reusedThisRun = false;

    if (pickedOtherLanguages.length < otherNeeded && otherLanguages.length > pickedOtherLanguages.length) {
      var alreadyPicked = {};
      for (var i = 0; i < pickedOtherLanguages.length; i++) {
        alreadyPicked[pickedOtherLanguages[i]] = true;
      }

      var refillPool = otherLanguages.filter(function (lang) {
        return !alreadyPicked[lang] && tracking._currentRunUsedLanguages.indexOf(lang) === -1;
      });
      var globallyUsedRefillPool = refillPool.filter(function (lang) {
        return tracking.usedLanguagesGlobalCycle.indexOf(lang) !== -1;
      });
      pickedOtherLanguages = pickedOtherLanguages.concat(shuffle(globallyUsedRefillPool).slice(0, otherNeeded - pickedOtherLanguages.length));
      if (globallyUsedRefillPool.length) reusedFreshness = true;
    }

    if (pickedOtherLanguages.length < otherNeeded && otherLanguages.length > pickedOtherLanguages.length) {
      var alreadyPickedAgain = {};
      for (var j = 0; j < pickedOtherLanguages.length; j++) {
        alreadyPickedAgain[pickedOtherLanguages[j]] = true;
      }
      var finalRefillPool = otherLanguages.filter(function (lang) { return !alreadyPickedAgain[lang]; });
      pickedOtherLanguages = pickedOtherLanguages.concat(shuffle(finalRefillPool).slice(0, otherNeeded - pickedOtherLanguages.length));
      if (finalRefillPool.length) reusedThisRun = true;
    }

    if (categoryName === 'Save the Earth' && pickedOtherLanguages.length < otherNeeded) {
      var selectedForReuse = ['English'].concat(pickedOtherLanguages);
      var reuseMap = getSaveTheEarthReuseMap(chosenTitle, selectedForReuse);
      Object.keys(reuseMap).forEach(function (lang) {
        if (pickedOtherLanguages.length >= otherNeeded) return;
        if (languageToFile[lang]) return;
        languageToFile[lang] = reuseMap[lang];
        pickedOtherLanguages.push(lang);
      });
    }

    if (reusedFreshness) {
      warningMessages.push(categoryName + ': not enough globally fresh non-English languages were available for the chosen title, so languages from earlier runs were reused.');
    }

    if (reusedThisRun) {
      warningMessages.push(categoryName + ': not enough unused non-English languages were available for the chosen title, so some languages already used earlier in this import had to be reused.');
    }

    if (categoryName === 'Save the Earth' && pickedOtherLanguages.length > otherNeeded) {
      pickedOtherLanguages = pickedOtherLanguages.slice(0, otherNeeded);
    }

    var selectedLanguages = ['English'].concat(pickedOtherLanguages);
    var selectedFiles = [languageToFile.English];
    pickedOtherLanguages.forEach(function (lang) { selectedFiles.push(languageToFile[lang]); });

    if (!scanResult.isFlatSingleTitle && categoryTracking.usedEnglishTitlesCycle.indexOf(chosenTitle) === -1) {
      categoryTracking.usedEnglishTitlesCycle.push(chosenTitle);
    }

    pickedOtherLanguages.forEach(function (lang) {
      if (tracking._currentRunUsedLanguages.indexOf(lang) === -1) {
        tracking._currentRunUsedLanguages.push(lang);
      }
      if (tracking.usedLanguagesGlobalCycle.indexOf(lang) === -1) {
        tracking.usedLanguagesGlobalCycle.push(lang);
      }
    });

    var totalAvailableCount = 1 + Object.keys(languageToFile).filter(function (lang) { return lang !== 'English'; }).length;
    if (requestedCount > totalAvailableCount) {
      if (categoryName === 'Save the Earth') {
        warningMessages.push(categoryName + ': requested ' + requestedCount + ' slides, but only ' + totalAvailableCount + ' distinct language versions were available after exact match, title-family fallback, and reused-language fallback.');
      } else {
        warningMessages.push(categoryName + ': requested ' + requestedCount + ' slides, but only ' + totalAvailableCount + ' matching language versions exist for the chosen title.');
      }
    }

    return {
      chosenTitle: chosenTitle,
      chosenTitleDisplay: chosenTitleDisplay,
      selectedLanguages: selectedLanguages,
      selectedFiles: selectedFiles,
      warning: warningMessages.join(' '),
      tracking: tracking
    };
  }

  function maybeResetGlobalLanguageCycle(tracking, categoryDataList) {
    var union = {};
    categoryDataList.forEach(function (categoryData) {
      var titlesMap = categoryData.scanResult.titlesMap;
      Object.keys(titlesMap).forEach(function (title) {
        Object.keys(titlesMap[title]).forEach(function (lang) {
          if (lang !== 'English') union[lang] = true;
        });
      });
    });
    var allNonEnglishLanguages = Object.keys(union);
    if (!allNonEnglishLanguages.length) return;
    var allUsed = allNonEnglishLanguages.every(function (lang) {
      return tracking.usedLanguagesGlobalCycle.indexOf(lang) !== -1;
    });
    if (allUsed) {
      log('All global non-English languages across categories have now been used once. Starting a fresh language cycle on the next run.');
      tracking.usedLanguagesGlobalCycle = [];
    }
  }

  function callJsx(functionName, payload, callback) {
    var json = JSON.stringify(payload || {});
    var script = functionName + "('" + escapeForEval(json) + "')";
    cs.evalScript(script, callback);
  }

  function setSelectedRootFolder(folder) {
    if (!folder) return;
    selectedRootFolder = folder;
    folderPathInput.value = folder;
    persistSettings();
    setStatus('Selected folder: ' + folder);
  }

  function setRunning(on) {
    runBtn.disabled = !!on;
    runBtn.style.opacity = on ? '0.5' : '';
    runBtn.style.cursor  = on ? 'not-allowed' : '';
  }

  function handlePlacementResponse(result) {
    setRunning(false);
    try {
      var parsed = JSON.parse(result);
      if (!parsed.ok) {
        log('⚠ Placement error: ' + parsed.error);
        return;
      }
      var msg = '─────────────────────────────\nDone.';
      for (var i = 0; i < parsed.results.length; i++) {
        var r = parsed.results[i];
        msg += '\n' + r.categoryName + ' → track V' + r.targetTrack + ': imported ' + r.importedCount + ', placed ' + r.placedCount;
        if (typeof r.intervalSeconds === 'number') {
          msg += ', interval ' + r.intervalSeconds.toFixed(1) + 's';  // interval stays in seconds (short number)
        }
        if (r.warning) msg += ' | ⚠ ' + r.warning;
      }
      if (parsed.presetNote) msg += '\n' + parsed.presetNote;
      if (parsed.note) msg += '\n' + parsed.note;
      log(msg);
    } catch (e) {
      log('Raw response: ' + result);
    }
  }

  function previewAndPlaceBatches(options) {
    // If avoidFaces is ON, we do NOT export frames here — buildSafePlacementPlan
    // will export frames one at a time per slide as it steps through timecodes.
    // exportFrames: false here keeps the initial preview fast (just calculates timecodes).
    var initialAnalysisDir = '';
    callJsx('newPeaceMakerPreviewPlacementFrames', {
      batches: options.batches,
      targetTrack: options.targetTrack,
      ignoreV1: options.ignoreV1,
      analysisDir: initialAnalysisDir,
      exportFrames: false
    }, function (previewResult) {
      var previewParsed;
      try {
        previewParsed = JSON.parse(previewResult);
      } catch (e) {
        setStatus('Preview failed: ' + previewResult);
        setRunning(false);
        return;
      }

      if (!previewParsed.ok) {
        setStatus('Preview failed: ' + previewParsed.error);
        setRunning(false);
        return;
      }

      function placeResolvedPlan(resolvedPlan, info) {
        if (info && info.movedCount) {
          log('Face/head avoidance moved ' + info.movedCount + ' slides in time to find safer moments.');
        }
        if (info && info.unsafeFallbackCount) {
          log('Face/head avoidance could not find a fully clear moment for ' + info.unsafeFallbackCount + ' slides, so it used the least-bad nearby time.');
        }

        callJsx('newPeaceMakerImportAndPlaceMulti', {
          batches: options.batches,
          requestedCount: options.requestedCount,
          rootFolderName: options.rootFolderName,
          targetTrack: options.targetTrack,
          ignoreV1: options.ignoreV1,
          slideAnchor: options.slideAnchor,
          placementPlan: resolvedPlan
        }, handlePlacementResponse);
      }

      if (options.avoidFaces) {
        buildSafePlacementPlan(options, previewParsed.placementPlan, function (err, resolvedPlan, info) {
          if (err || !resolvedPlan) {
            setStatus('Safe placement failed: ' + (err ? err.message : 'Unknown error'));
            setRunning(false);
            return;
          }
          placeResolvedPlan(resolvedPlan, info || {});
        });
        return;
      }

      // No face avoidance — use JSX positions directly, but first repair any
      // V1-boundary overlaps that arise from the compressed-time mapping in JSX
      // (a slide placed at t=range.end can overlap the next V1 clip by its duration).
      if (options.ignoreV1 && previewParsed.placementPlan) {
        var plan0 = previewParsed.placementPlan;
        var v1Blocked = Array.isArray(plan0.blockedV1Ranges) ? plan0.blockedV1Ranges : [];
        var wEnd0 = (typeof plan0.placementWindowEndSeconds === 'number' && plan0.placementWindowEndSeconds > 0)
          ? plan0.placementWindowEndSeconds
          : (typeof plan0.usedTimelineLengthSeconds === 'number' ? plan0.usedTimelineLengthSeconds : 0);
        if (v1Blocked.length > 0) {
          (plan0.placements || []).forEach(function (p, pi) {
            var dur = p.clipDurationSeconds || 9;
            if (overlapsBlockedRange(p.startSeconds, dur, v1Blocked)) {
              var adj = nextAvailableStart(p.startSeconds, dur, v1Blocked);
              if (wEnd0 <= 0 || adj + dur <= wEnd0) {
                log('Slide ' + (pi + 1) + ': V1 boundary overlap at ' + secToMS(p.startSeconds) + ' → adjusted to ' + secToMS(adj) + '.');
                p.startSeconds = adj;
              }
            }
          });
        }
      }

      placeResolvedPlan(previewParsed.placementPlan, { movedCount: 0, unsafeFallbackCount: 0 });
    });
  }

  browseBtn.addEventListener('click', function () {
    try {
      if (window.cep && window.cep.fs && typeof window.cep.fs.showOpenDialogEx === 'function') {
        var result = window.cep.fs.showOpenDialogEx(false, true, 'Select the root folder that contains your slide files');
        if (result && result.data && result.data.length) {
          setSelectedRootFolder(result.data[0]);
          return;
        }
      }
    } catch (e) {
      log('CEP folder dialog failed, falling back to HTML picker. ' + e.message);
    }

    folderPicker.click();
  });

  folderPicker.addEventListener('change', function (evt) {
    var files = evt.target.files;
    if (!files || !files.length) return;

    var rel = files[0].webkitRelativePath || '';
    var topFolder = rel.split('/')[0];
    var firstAbsolute = files[0].path;
    selectedRootFolder = path.dirname(firstAbsolute);

    while (path.basename(selectedRootFolder) !== topFolder && selectedRootFolder !== path.dirname(selectedRootFolder)) {
      selectedRootFolder = path.dirname(selectedRootFolder);
    }

    setSelectedRootFolder(selectedRootFolder);
  });


  slideCountInput.addEventListener('change', persistSettings);
  targetTrackInput.addEventListener('change', persistSettings);
  ignoreV1Input.addEventListener('change', persistSettings);
  slideAnchorInput.addEventListener('change', persistSettings);
  avoidFacesInput.addEventListener('change', persistSettings);
  installUpdateBtn.addEventListener('click', installLatestUpdate);
  window.addEventListener('beforeunload', persistSettings);

  runBtn.addEventListener('click', function () {
    try {
      setStatus('Working...');
      if (!selectedRootFolder) {
        setStatus('Please choose the root folder that contains your slide files first.');
        return;
      }

      var requestedCount = parseInt(slideCountInput.value, 10);
      var targetTrack = parseInt(targetTrackInput.value, 10);
      var ignoreV1 = !!ignoreV1Input.checked;
      var slideAnchor = String(slideAnchorInput.value || 'top-right');
      var avoidFaces = !!avoidFacesInput.checked;

      if (!requestedCount || requestedCount < 1) {
        setStatus('Please enter a valid number of slides.');
        return;
      }
      if (!targetTrack || targetTrack < 1) {
        setStatus('Please enter a valid target video track number.');
        return;
      }

      setRunning(true);
      persistSettings();
      var tracking = loadTracking();
      tracking._currentRunUsedLanguages = [];
      var categoryDataList = scanAllCategories(selectedRootFolder);
      if (!categoryDataList.length) {
        setStatus('No usable slide files were found under the selected root folder.');
        return;
      }

      log('Ignored folder: AFTERCODECS HAP ALPHA');
      var batches = [];
      var titleSummary = [];
      var languageSummary = [];

      categoryDataList.forEach(function (categoryData, index) {
        if (!categoryData.scanResult.titlesMap || !Object.keys(categoryData.scanResult.titlesMap).length) {
          throw new Error('Category "' + categoryData.name + '" has no usable slide files.');
        }
        var hasEnglish = Object.keys(categoryData.scanResult.titlesMap).some(function (title) {
          return !!categoryData.scanResult.titlesMap[title].English;
        });
        if (!hasEnglish) {
          throw new Error('Category "' + categoryData.name + '" has no usable English file.');
        }
        var batch = chooseBatchForCategory(categoryData.name, categoryData.scanResult, requestedCount, tracking);
        batches.push({
          categoryName: categoryData.name,
          files: batch.selectedFiles,
          languages: batch.selectedLanguages,
          languageDetails: batch.selectedLanguages.map(function (lang) {
            return {
              name: lang,
              isEnglish: lang === 'English'
            };
          }),
          title: batch.chosenTitleDisplay,
          targetTrack: targetTrack,
          warning: batch.warning || ''
        });
        titleSummary.push(categoryData.name + ': ' + batch.chosenTitleDisplay);
        languageSummary.push(categoryData.name + ': ' + batch.selectedLanguages.join(', '));
        if (batch.warning) log('Warning: ' + batch.warning);
      });

      maybeResetGlobalLanguageCycle(tracking, categoryDataList);
      delete tracking._currentRunUsedLanguages;
      saveTracking(tracking);

      chosenTitleEl.textContent = titleSummary.join('\n');
      chosenLanguagesEl.textContent = languageSummary.join('\n');

      if (avoidFaces) {
        log('Analyzing visible sequence frames for heads/faces before placement...');
      }
      previewAndPlaceBatches({
        batches: batches,
        requestedCount: requestedCount,
        rootFolderName: path.basename(selectedRootFolder),
        targetTrack: targetTrack,
        ignoreV1: ignoreV1,
        slideAnchor: slideAnchor,
        avoidFaces: avoidFaces
      });
    } catch (err) {
      setStatus('Error: ' + err.message);
      setRunning(false);
    }
  });

  extensionRoot = resolveExtensionRoot();
  manifestPath = extensionRoot ? path.join(extensionRoot, 'CSXS', 'manifest.xml') : path.join(path.resolve(__dirname, '..'), 'CSXS', 'manifest.xml');
  restoreSettings();
  initFaceApi();    // Load TinyFaceDetector model (async — falls back to pixel analysis until ready)
  initOCRAD();  // Load OCRAD.js OCR — synchronous, no Workers, no SharedArrayBuffer
  updateState.installedVersion = readManifestVersion(manifestPath) || '';
  updateState.latestVersion = loadTracking().settings.lastAvailableVersion || '';
  var updateInstallStatus = loadUpdateInstallStatus();
  var pendingUpdateInfo = getPendingUpdateInfo();
  var justCompletedUpdate = false;
  if (updateInstallStatus && updateInstallStatus.state === 'success' && compareVersions(updateState.installedVersion, updateInstallStatus.version) >= 0) {
    justCompletedUpdate = true;
    setUpdateStatus('Update installed successfully.');
      showUpdateModal(
        'What Is New',
        buildUpdateNotesMessage({
          tag_name: pendingUpdateInfo.version || updateInstallStatus.version || updateState.installedVersion,
          name: pendingUpdateInfo.name || updateInstallStatus.version || updateState.installedVersion,
          body: pendingUpdateInfo.notes || ''
        }, 'The update was installed successfully.', { popupSummary: true }),
        { okText: 'OK' }
      ).then(function () {
      clearUpdateInstallStatus();
      clearPendingUpdateInfo();
    });
    updateInstallStatus = null;
  }

  if (updateInstallStatus && updateInstallStatus.state === 'failed') {
    setUpdateStatus('Previous update failed: ' + (updateInstallStatus.message || 'Unknown error') + '.');
  } else if (updateInstallStatus && (updateInstallStatus.state === 'staged' || updateInstallStatus.state === 'pending')) {
    setUpdateStatus('A staged update to version ' + (updateInstallStatus.version || '?') + ' is still pending. Close Premiere Pro fully and wait a few seconds before reopening. If Windows asked for permission, accept the prompt.');
  } else if (!justCompletedUpdate) {
    setUpdateStatus(updateRepo ? 'Ready to check for updates.' : 'No GitHub update source is configured.');
  }
  setUpdateUiState();
  if (updateRepo) {
    checkForUpdates({ silent: true });
  }
})();
