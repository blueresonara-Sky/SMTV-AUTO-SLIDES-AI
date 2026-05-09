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
  var packPerCategoryInput = document.getElementById('packPerCategory');
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

  // ── Quan-Yin & Max tab DOM refs ──────────────────────────────────────────
  var qymQYFolderInput     = document.getElementById('qymQYFolder');
  var qymMaxFolderInput    = document.getElementById('qymMaxFolder');
  var qymQYBrowseBtn       = document.getElementById('qymQYBrowse');
  var qymMaxBrowseBtn      = document.getElementById('qymMaxBrowse');
  var qymQYPickerInput     = document.getElementById('qymQYPicker');
  var qymMaxPickerInput    = document.getElementById('qymMaxPicker');
  var qymMaxCountInput     = document.getElementById('qymMaxCount');
  var qymTargetTrackInput  = document.getElementById('qymTargetTrack');
  var qymFadeDurationInput = document.getElementById('qymFadeDuration');
  var qymAvoidFacesInput   = document.getElementById('qymAvoidFaces');
  var qymKeepDebugInput    = document.getElementById('qymKeepDebug');
  var qymRunBtn            = document.getElementById('qymRunBtn');
  var qymStatusEl          = document.getElementById('qymStatus');
  // Slide-type radio group (values: 'both', 'qy-only', 'max-only')
  var qymSlideTypeRadios   = document.getElementsByName('qymSlideType');
  // No-In/Out confirmation modal
  var qymConfirmModalEl    = document.getElementById('qymConfirmModal');
  var qymConfirmOkBtn      = document.getElementById('qymConfirmOkBtn');
  var qymConfirmCancelBtn  = document.getElementById('qymConfirmCancelBtn');

  function getQymSlideType() {
    if (!qymSlideTypeRadios) return 'both';
    for (var i = 0; i < qymSlideTypeRadios.length; i++) {
      if (qymSlideTypeRadios[i].checked) return qymSlideTypeRadios[i].value;
    }
    return 'both';
  }
  function setQymSlideType(value) {
    if (!qymSlideTypeRadios) return;
    for (var i = 0; i < qymSlideTypeRadios.length; i++) {
      qymSlideTypeRadios[i].checked = (qymSlideTypeRadios[i].value === value);
    }
  }
  // Modal helper — opens the no-in/out modal and calls cb(true|false) when
  // the user clicks Continue / Cancel.
  function showQymNoInOutConfirm(cb) {
    if (!qymConfirmModalEl) { cb(true); return; }
    qymConfirmModalEl.classList.add('is-open');
    var onOk = function () { cleanup(); cb(true); };
    var onCancel = function () { cleanup(); cb(false); };
    function cleanup() {
      qymConfirmModalEl.classList.remove('is-open');
      if (qymConfirmOkBtn)     qymConfirmOkBtn.removeEventListener('click', onOk);
      if (qymConfirmCancelBtn) qymConfirmCancelBtn.removeEventListener('click', onCancel);
    }
    if (qymConfirmOkBtn)     qymConfirmOkBtn.addEventListener('click', onOk);
    if (qymConfirmCancelBtn) qymConfirmCancelBtn.addEventListener('click', onCancel);
  }

  var selectedQYFolder  = '';
  var selectedMaxFolder = '';
  var qymRunning        = false;

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
  var faceApiReady  = false;
  var cocoSsdReady  = false;
  var cocoSsdModel  = null;
  var cocoSsdInitError = '';

  // ── Loading banner — shown until all detection models finish initialising
  var loadingBannerEl = null;
  var loadingTextEl   = null;
  var modelLoadState  = {
    faceApi:  'loading',   // 'loading' | 'ready' | 'failed'
    cocoSsd:  'loading',
    ocrad:    'loading'
  };
  function _updateLoadingBanner() {
    if (!loadingBannerEl) {
      loadingBannerEl = document.getElementById('loadingBanner');
      loadingTextEl   = document.getElementById('loadingText');
    }
    if (!loadingBannerEl || !loadingTextEl) return;
    var appShell = document.querySelector('.app-shell');

    var stillLoading = [];
    if (modelLoadState.faceApi === 'loading') stillLoading.push('face-api (Slides tab)');
    if (modelLoadState.cocoSsd === 'loading') stillLoading.push('coco-ssd (QYM tab)');
    if (modelLoadState.ocrad   === 'loading') stillLoading.push('OCRAD');
    if (stillLoading.length === 0) {
      // Everything done — hide banner and un-fade the rest of the panel.
      loadingTextEl.textContent = 'All detection models ready.';
      loadingBannerEl.style.background = '#1f4f31';
      loadingBannerEl.style.borderColor = '#3d8e57';
      loadingBannerEl.style.color = '#ffffff';
      var spin = loadingBannerEl.querySelector('.loading-spinner');
      if (spin) spin.style.display = 'none';
      // Remove the fade class so the panel becomes fully interactive.
      if (appShell) appShell.classList.remove('models-loading');
      setTimeout(function () { loadingBannerEl.classList.add('hidden'); }, 1200);
    } else {
      loadingTextEl.textContent = 'Loading detection models — please wait… (' + stillLoading.join(', ') + ')';
      // Make sure the fade class is on while we're still loading.
      if (appShell) appShell.classList.add('models-loading');
    }
  }
  function _markModelReady(name)  { modelLoadState[name] = 'ready';  _updateLoadingBanner(); _updateRunButtonsForLoading(); }
  function _markModelFailed(name) { modelLoadState[name] = 'failed'; _updateLoadingBanner(); _updateRunButtonsForLoading(); }
  function _updateRunButtonsForLoading() {
    // Disable BOTH run buttons until at least the relevant detector for each
    // tab is settled (ready or failed). Prevents a click landing on coco-ssd
    // before the model file finishes loading.
    var qymWaiting    = (modelLoadState.cocoSsd === 'loading');
    var slidesWaiting = (modelLoadState.faceApi === 'loading');
    if (qymRunBtn && !qymRunning) {
      qymRunBtn.disabled = qymWaiting;
      if (qymWaiting) qymRunBtn.title = 'Waiting for coco-ssd model to finish loading…';
      else            qymRunBtn.title = '';
    }
    var runBtnEl = document.getElementById('runBtn');
    if (runBtnEl) {
      runBtnEl.disabled = slidesWaiting;
      if (slidesWaiting) runBtnEl.title = 'Waiting for face-api model to finish loading…';
      else               runBtnEl.title = '';
    }
  }
  var faceApiInitError = null;

  // OCR text-detection state (OCRAD.js)
  var ocradReady = false;

  function defaultTracking() {
    return {
      categories: {},
      usedLanguagesGlobalCycle: [],
      ignoredFolders: ['AFTERCODECS HAP ALPHA'],
      // ── Quan-Yin & Max cycle tracking ─────────────────────
      quanYinCycle: {
        nextEnglish:   'A',  // toggles A → B → A each run
        usedNonEnglish: []   // language names used this cycle; reset when all consumed
      },
      smtvMaxCycle: {
        usedNonEnglish: []   // language codes used this cycle
      },
      settings: {
        rootFolder: '',
        slideCount: 6,
        targetTrack: 9,
        ignoreV1: false,
        slideAnchor: 'top-right',
        avoidFaces: true,
        packPerCategory: false,
        lastUpdateCheckAt: '',
        lastAvailableVersion: '',
        pendingUpdateVersion: '',
        pendingUpdateName: '',
        pendingUpdateNotes: '',
        // ── QYM settings ──────────────────────────────────────
        qymQuanYinFolder: '',
        qymSmtvMaxFolder: '',
        qymMaxCount:       4,
        qymTargetTrack:   10,
        qymFadeDuration:   0.3,
        qymAvoidFaces:    true,
        qymSlideType:    'both',  // 'both' | 'qy-only' | 'max-only'
        qymKeepDebug:    false,   // when true, save annotated frames to ~/qym-debug/
      }
    };
  }

  function log(msg) {
    statusEl.textContent += '\n' + msg;
    statusEl.scrollTop = statusEl.scrollHeight;
  }

  // Format seconds as m:ss:ff at 29.97 fps  e.g. 203.7 → "3:23:21"
  // Format real-time seconds as drop-frame timecode (29.97 fps DF), matching
  // Premiere's default display. Earlier this function ignored DF and produced
  // labels that drifted ~2 frames per minute from Premiere's TC — by 9 minutes
  // the drift was ~16 frames (~0.5s), making it impossible to map log entries
  // back to actual sequence positions.
  //
  // SMPTE drop-frame algorithm:
  //   - Drop 2 frames per minute, EXCEPT every 10th minute.
  //   - This keeps DF timecode aligned with wall-clock time.
  function secToMS(s) {
    var fps         = 29.97;
    var nominalFps  = 30;
    var dropPerMin  = 2;
    var neg         = s < 0;
    var abs         = Math.abs(s);
    var frameNum    = Math.round(abs * fps);

    // Frames per 10-minute block in drop-frame: 9 minutes drop 2 each (=18) +
    // 1 minute (10th) drops 0 → total drops per 10 min = 18.
    // Nominal frames per 10 min = 18000; DF = 18000 - 18 = 17982.
    var framesPer10Min  = nominalFps * 60 * 10 - dropPerMin * 9;       // 17982
    var framesPerMinute = nominalFps * 60 - dropPerMin;                // 1798

    var d = Math.floor(frameNum / framesPer10Min);
    var m = frameNum % framesPer10Min;
    if (m > dropPerMin) {
      frameNum = frameNum + dropPerMin * 9 * d + dropPerMin * Math.floor((m - dropPerMin) / framesPerMinute);
    } else {
      frameNum = frameNum + dropPerMin * 9 * d;
    }

    var ff   = frameNum % nominalFps;
    var ss   = Math.floor(frameNum / nominalFps) % 60;
    var mm   = Math.floor(frameNum / (nominalFps * 60));
    var pad  = function (n) { return n < 10 ? '0' + n : '' + n; };
    return (neg ? '-' : '') + mm + ':' + pad(ss) + ':' + pad(ff);
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
      // Quan-Yin / Max cycle migration
      parsed.quanYinCycle = parsed.quanYinCycle || {};
      if (typeof parsed.quanYinCycle.nextEnglish    !== 'string') parsed.quanYinCycle.nextEnglish    = 'A';
      if (!Array.isArray(parsed.quanYinCycle.usedNonEnglish))     parsed.quanYinCycle.usedNonEnglish  = [];
      parsed.smtvMaxCycle = parsed.smtvMaxCycle || {};
      if (!Array.isArray(parsed.smtvMaxCycle.usedNonEnglish))     parsed.smtvMaxCycle.usedNonEnglish  = [];
      parsed.settings = parsed.settings || base.settings;
      if (typeof parsed.settings.slideCount === 'undefined') parsed.settings.slideCount = base.settings.slideCount;
      if (typeof parsed.settings.targetTrack === 'undefined') parsed.settings.targetTrack = base.settings.targetTrack;
      if (typeof parsed.settings.rootFolder === 'undefined') parsed.settings.rootFolder = base.settings.rootFolder;
      if (typeof parsed.settings.ignoreV1 === 'undefined') parsed.settings.ignoreV1 = base.settings.ignoreV1;
      if (typeof parsed.settings.slideAnchor === 'undefined') parsed.settings.slideAnchor = base.settings.slideAnchor;
      if (typeof parsed.settings.avoidFaces === 'undefined') parsed.settings.avoidFaces = base.settings.avoidFaces;
      if (typeof parsed.settings.packPerCategory === 'undefined') parsed.settings.packPerCategory = base.settings.packPerCategory;
      if (typeof parsed.settings.lastUpdateCheckAt === 'undefined') parsed.settings.lastUpdateCheckAt = base.settings.lastUpdateCheckAt;
      if (typeof parsed.settings.lastAvailableVersion === 'undefined') parsed.settings.lastAvailableVersion = base.settings.lastAvailableVersion;
      if (typeof parsed.settings.pendingUpdateVersion === 'undefined') parsed.settings.pendingUpdateVersion = base.settings.pendingUpdateVersion;
      if (typeof parsed.settings.pendingUpdateName === 'undefined') parsed.settings.pendingUpdateName = base.settings.pendingUpdateName;
      if (typeof parsed.settings.pendingUpdateNotes === 'undefined') parsed.settings.pendingUpdateNotes = base.settings.pendingUpdateNotes;
      // QYM settings migration
      if (typeof parsed.settings.qymQuanYinFolder === 'undefined') parsed.settings.qymQuanYinFolder = base.settings.qymQuanYinFolder;
      if (typeof parsed.settings.qymSmtvMaxFolder === 'undefined') parsed.settings.qymSmtvMaxFolder = base.settings.qymSmtvMaxFolder;
      if (typeof parsed.settings.qymMaxCount      === 'undefined') parsed.settings.qymMaxCount      = base.settings.qymMaxCount;
      if (typeof parsed.settings.qymTargetTrack   === 'undefined') parsed.settings.qymTargetTrack   = base.settings.qymTargetTrack;
      if (typeof parsed.settings.qymFadeDuration  === 'undefined') parsed.settings.qymFadeDuration  = base.settings.qymFadeDuration;
      if (typeof parsed.settings.qymAvoidFaces    === 'undefined') parsed.settings.qymAvoidFaces    = base.settings.qymAvoidFaces;
      if (typeof parsed.settings.qymSlideType     === 'undefined') parsed.settings.qymSlideType     = base.settings.qymSlideType;
      if (typeof parsed.settings.qymKeepDebug     === 'undefined') parsed.settings.qymKeepDebug     = base.settings.qymKeepDebug;
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
    tracking.settings.packPerCategory = !!(packPerCategoryInput && packPerCategoryInput.checked);
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
    if (packPerCategoryInput) packPerCategoryInput.checked = !!tracking.settings.packPerCategory;
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

  // ── Face-detection region override ────────────────────────────────────
  // When set, runFaceApiAnalysis / scoreAnchorRegion / detectTextInRegion use
  // these regions instead of their built-in (Slides-tab badge-strip) defaults.
  // The QYM tab sets this in buildSafePlacementPlanQYM because Quan-Yin/SMTV-Max
  // ad cards are MUCH larger and lower than the original tab's tiny top badges.
  // Format: { 'top-left': {x1,y1,x2,y2}, 'top-right': {x1,y1,x2,y2} }  (fractions)
  var _qymFaceRegionOverride = null;

  // Returns the region to use for a given anchor. If override is active,
  // it wins; otherwise returns null and the caller uses its own default.
  function _getOverrideRegion(anchorKey) {
    if (!_qymFaceRegionOverride) return null;
    var r = _qymFaceRegionOverride[anchorKey];
    return r || null;
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
    // QYM override (covers a much larger / lower footprint)
    var ov = _getOverrideRegion(anchorKey);
    var region;
    if (ov) {
      region = { x: ov.x1, y: ov.y1, w: ov.x2 - ov.x1, h: ov.y2 - ov.y1 };
    } else {
      region = regions[anchorKey] || regions['top-right'];
    }
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
    // QYM override
    var ovT = _getOverrideRegion(anchorKey);
    var reg;
    if (ovT) {
      reg = { x: ovT.x1, y: ovT.y1, w: ovT.x2 - ovT.x1, h: ovT.y2 - ovT.y1 };
    } else {
      reg = regionDefs[anchorKey];
    }
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
      _markModelFailed('faceApi');
      return;
    }

    var modelsDir = extensionRoot ? path.join(extensionRoot, 'client', 'models') : '';
    if (!modelsDir || !fs.existsSync(modelsDir)) {
      log('[face-api] Models folder not found at "' + modelsDir + '" — pixel-based fallback active.');
      _markModelFailed('faceApi');
      return;
    }

    var manifestFile = path.join(modelsDir, 'tiny_face_detector_model-weights_manifest.json');
    var binFile      = path.join(modelsDir, 'tiny_face_detector_model.bin');
    if (!fs.existsSync(manifestFile) || !fs.existsSync(binFile)) {
      log('[face-api] Model files missing in "' + modelsDir + '" — pixel-based fallback active.');
      _markModelFailed('faceApi');
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
        _markModelReady('faceApi');
      })
      .catch(function (err) {
        faceApiInitError = err && err.message ? err.message : String(err);
        log('[face-api] Model load failed: ' + faceApiInitError + ' — pixel-based fallback active.');
        _markModelFailed('faceApi');
      });
  }

  // ── COCO-SSD person detector (TensorFlow.js) ───────────────────────────────
  //
  // Used by the QYM tab ONLY. Detects whole people (head + body, any pose,
  // hijab, profile, prostrating, etc.) — much more robust than face-api's
  // frontal-face-only detection.
  //
  // Model files expected in client/lib/:
  //   model.json + group1-shard1of5 .. group1-shard5of5 (~5MB total)
  //
  // The Slides tab still uses face-api; this is purely a QYM enhancement.

  function initCocoSsd() {
    var cocoGlobal = (typeof cocoSsd !== 'undefined') ? cocoSsd
      : (typeof window !== 'undefined' && window.cocoSsd) ? window.cocoSsd
      : null;

    if (!cocoGlobal) {
      logQYM('[coco-ssd] Library not loaded — QYM person detection unavailable.');
      _markModelFailed('cocoSsd');
      return;
    }
    if (typeof tf === 'undefined' && (typeof window === 'undefined' || !window.tf)) {
      logQYM('[coco-ssd] TensorFlow.js not loaded — QYM person detection unavailable.');
      _markModelFailed('cocoSsd');
      return;
    }

    var libDir = extensionRoot ? path.join(extensionRoot, 'client', 'lib') : '';
    var modelFile = path.join(libDir, 'model.json');
    if (!fs.existsSync(modelFile)) {
      logQYM('[coco-ssd] model.json not found at "' + modelFile + '" — QYM person detection unavailable.');
      _markModelFailed('cocoSsd');
      return;
    }

    // file:// URI for the Chromium/CEF runtime
    var modelUri;
    if (process.platform === 'win32') {
      modelUri = 'file:///' + modelFile.replace(/\\/g, '/').replace(/ /g, '%20');
    } else {
      modelUri = 'file://' + modelFile.replace(/ /g, '%20');
    }

    logQYM('[coco-ssd] Loading model from: ' + modelUri);

    // cocoSsd.load supports a `modelUrl` option to load a custom local model.
    cocoGlobal.load({ modelUrl: modelUri })
      .then(function (model) {
        cocoSsdModel = model;
        cocoSsdReady = true;
        logQYM('[coco-ssd] ✓ Model ready — person detection active for QYM.');
        _markModelReady('cocoSsd');
      })
      .catch(function (err) {
        cocoSsdInitError = err && err.message ? err.message : String(err);
        logQYM('[coco-ssd] Model load failed: ' + cocoSsdInitError);
        _markModelFailed('cocoSsd');
      });
  }

  // Run coco-ssd on a frame canvas and return all 'person' bounding boxes
  // normalised to 0–1 frame fractions. Resolves to [] on any failure or
  // when the model isn't ready.
  function detectPersonsAsync(canvas) {
    return new Promise(function (resolve) {
      if (!cocoSsdReady || !cocoSsdModel) { resolve([]); return; }
      try {
        cocoSsdModel.detect(canvas).then(function (predictions) {
          var w = canvas.width, h = canvas.height;
          var people = [];
          for (var i = 0; i < predictions.length; i++) {
            var p = predictions[i];
            if (p.class !== 'person') continue;
            var bx = p.bbox[0], by = p.bbox[1], bw = p.bbox[2], bh = p.bbox[3];
            people.push({
              x1:    bx / w,
              y1:    by / h,
              x2:    (bx + bw) / w,
              y2:    (by + bh) / h,
              score: p.score
            });
          }
          resolve(people);
        }).catch(function () { resolve([]); });
      } catch (e) { resolve([]); }
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
      _markModelFailed('ocrad');
      return;
    }
    ocradReady = true;
    log('[text-detect] ✓ OCRAD ready — OCR text detection active.');
    _markModelReady('ocrad');
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

  // ── QYM-specific OCRAD text detection ─────────────────────────────────────
  // Runs OCRAD on a tightly-margined zone around the QYM slide region to
  // catch broadcast captions baked into the rendered frame (e.g. "Words of
  // Wisdom" tabs, presenter ribbons). When real text is found in this zone,
  // the candidate is rejected so the slide doesn't sit on top of a caption.
  //
  // Asymmetric margins (per user spec):
  //   top    = 0%   tight — captions ABOVE the slide are common; we don't
  //                 want to flag every clip just because of them.
  //   right  = 0%   slide hugs the frame edge already.
  //   left   = +10% catch text whose bbox extends in from the left.
  //   bottom = +10% catch presenter ribbons that intrude up from below.
  //
  // Returns: { detected: bool, text: '...' (≤ 80 chars) }
  function _qymDetectTextInRegion(canvas) {
    var ocradFn = (typeof OCRAD !== 'undefined') ? OCRAD
      : (typeof window !== 'undefined' && window.OCRAD) ? window.OCRAD
      : null;
    if (!ocradFn) return { detected: false, text: '' };

    var slideReg = (_qymFaceRegionOverride && _qymFaceRegionOverride['top-right'])
      ? _qymFaceRegionOverride['top-right']
      : null;
    if (!slideReg) return { detected: false, text: '' };

    var marginTop    = 0.00;
    var marginBottom = 0.10;
    var marginLeft   = 0.10;
    var marginRight  = 0.00;

    var rx1 = Math.max(0, slideReg.x1 - marginLeft);
    var ry1 = Math.max(0, slideReg.y1 - marginTop);
    var rx2 = Math.min(1, slideReg.x2 + marginRight);
    var ry2 = Math.min(1, slideReg.y2 + marginBottom);

    var W = canvas.width, H = canvas.height;
    var rx = Math.floor(W * rx1);
    var ry = Math.floor(H * ry1);
    var rw = Math.min(Math.floor(W * (rx2 - rx1)), W - rx);
    var rh = Math.min(Math.floor(H * (ry2 - ry1)), H - ry);
    if (rw <= 4 || rh <= 4) return { detected: false, text: '' };

    // 2× upscale — broadcast captions are often small; OCRAD reads them better at higher resolution.
    var scale = 2;
    var crop = document.createElement('canvas');
    crop.width  = rw * scale;
    crop.height = rh * scale;
    var cropCtx = crop.getContext('2d');
    cropCtx.drawImage(canvas, rx, ry, rw, rh, 0, 0, crop.width, crop.height);

    var N = crop.width * crop.height;
    var imgData = cropCtx.getImageData(0, 0, crop.width, crop.height);
    var d = imgData.data;

    var lums = new Array(N);
    var pi;
    for (pi = 0; pi < N; pi++) {
      lums[pi] = 0.299 * d[pi * 4] + 0.587 * d[pi * 4 + 1] + 0.114 * d[pi * 4 + 2];
    }

    // Skip flat regions — uniform pixels can't contain legible text.
    var mean = 0;
    for (pi = 0; pi < N; pi++) mean += lums[pi];
    mean /= N;
    var variance = 0;
    for (pi = 0; pi < N; pi++) {
      var dv = lums[pi] - mean;
      variance += dv * dv;
    }
    variance /= N;
    if (variance < 300) return { detected: false, text: '' };

    // Otsu's auto-threshold
    var hist = new Array(256);
    var ti;
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

    // Try four binarizations to handle any text/background combo.
    // Match real words: 3+ consecutive Latin letters.
    var textRe = /[A-Za-z]{3,}/;
    var foundText = '';
    function tryOCRADAt(threshold, invert) {
      if (foundText) return;
      for (var pp = 0; pp < N; pp++) {
        var v = (lums[pp] > threshold) ? 255 : 0;
        if (invert) v = 255 - v;
        d[pp * 4] = d[pp * 4 + 1] = d[pp * 4 + 2] = v;
        d[pp * 4 + 3] = 255;
      }
      cropCtx.putImageData(imgData, 0, 0);
      try {
        var raw = ocradFn(crop);
        if (textRe.test(raw)) foundText = raw.replace(/\s+/g, ' ').trim();
      } catch (e) {}
    }
    tryOCRADAt(otsuT, true);   // dark text on light bg (OCRAD's preferred)
    tryOCRADAt(otsuT, false);  // light text on dark bg
    tryOCRADAt(200,    true);  // fixed bright-text isolation
    tryOCRADAt(200,    false); // fallback

    return { detected: !!foundText, text: foundText.slice(0, 80) };
  }

  // ── QYM graphic-overlay detector ────────────────────────────────────────────
  // Replaces OCRAD for the QYM tab. Instead of reading characters, it measures
  // the density of hard-edge pixels (very sharp gradient transitions) in the TR
  // region. Composited broadcast graphics — regardless of color, shape, or
  // whether they have a border — have pixel-perfect 1–2px edges that produce
  // gradient magnitudes far above anything camera-captured video generates
  // (camera optics + H.264 compression spread natural edges over 4–10px).
  //
  // Same region and margins as _qymDetectTextInRegion (left+10%, bottom+10%).
  // Returns: { detected: bool, text: string (density info for debug log) }
  function _qymDetectGraphicInRegion(canvas, slideRegOverride) {
    var slideReg = slideRegOverride
      || ((_qymFaceRegionOverride && _qymFaceRegionOverride['top-right'])
          ? _qymFaceRegionOverride['top-right'] : null);
    if (!slideReg) return { detected: false, text: '' };

    var rx1 = Math.max(0, slideReg.x1 - 0.10);
    var ry1 = Math.max(0, slideReg.y1 - 0.00);
    var rx2 = Math.min(1, slideReg.x2 + 0.00);
    var ry2 = Math.min(1, slideReg.y2 + 0.10);

    var W = canvas.width, H = canvas.height;
    var rx = Math.floor(W * rx1), ry = Math.floor(H * ry1);
    var rw = Math.min(Math.floor(W * (rx2 - rx1)), W - rx);
    var rh = Math.min(Math.floor(H * (ry2 - ry1)), H - ry);
    if (rw <= 4 || rh <= 4) return { detected: false, text: '' };

    var crop = document.createElement('canvas');
    crop.width  = rw;
    crop.height = rh;
    var cropCtx = crop.getContext('2d');
    cropCtx.drawImage(canvas, rx, ry, rw, rh, 0, 0, rw, rh);

    var imgData = cropCtx.getImageData(0, 0, rw, rh);
    var d = imgData.data;
    var CW = rw, CH = rh;

    var lums = new Float32Array(CW * CH);
    var pi;
    for (pi = 0; pi < CW * CH; pi++) {
      lums[pi] = 0.299 * d[pi * 4] + 0.587 * d[pi * 4 + 1] + 0.114 * d[pi * 4 + 2];
    }

    // Count pixels whose gradient magnitude exceeds the hard-edge threshold.
    // Composited graphics: 1-2px border → gradient 150-400+.
    // Natural video edges: softened by optics/codec → rarely exceed 60-80.
    var HARD_THRESHOLD = 100;
    var MIN_DENSITY    = 0.025;  // 2.5% of interior pixels must be hard-edge
    var hardCount = 0, totalInner = 0;
    var x, y, idx, dx, dy;
    for (y = 1; y < CH - 1; y++) {
      for (x = 1; x < CW - 1; x++) {
        idx = y * CW + x;
        dx = Math.abs(lums[idx + 1]  - lums[idx - 1]);
        dy = Math.abs(lums[idx + CW] - lums[idx - CW]);
        if (dx + dy > HARD_THRESHOLD) hardCount++;
        totalInner++;
      }
    }

    var density  = totalInner > 0 ? hardCount / totalInner : 0;
    var detected = density >= MIN_DENSITY;
    return {
      detected: detected,
      text: detected ? ('[graphic d=' + (density * 100).toFixed(1) + '%]') : ''
    };
  }

  // Returns a Promise that resolves to { 'top-left': ratio, 'top-right': ratio }
  // ratio is 0.0 (no text) or 0.9 (text detected by OCRAD), or 0–1 from pixel fallback.
  // Never rejects.
  function detectTextInFrameAsync(canvas) {
    return new Promise(function (resolve) {
      // QYM mode: skip OCRAD (slow, ~300–600ms per frame). The pixel-based
      // detector is plenty for our purpose — we just want to know if a text
      // overlay (e.g. ticker, lower-third) sits under the proposed slide
      // region; we don't need character recognition.
      if (ocradReady && !_qymFaceRegionOverride) {
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
    // QYM uses much larger / lower regions (centered slide cards, not tiny corner badges).
    if (_qymFaceRegionOverride) {
      if (_qymFaceRegionOverride['top-left'])  regionDefs['top-left']  = _qymFaceRegionOverride['top-left'];
      if (_qymFaceRegionOverride['top-right']) regionDefs['top-right'] = _qymFaceRegionOverride['top-right'];
    }

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
      // Lower threshold in QYM mode — region is large and may contain
      // profile/turned/hijab heads that the default 0.4 threshold would miss.
      var detThreshold = _qymFaceRegionOverride ? 0.25 : 0.4;
      detectorOptions = new faceApiGlobal.TinyFaceDetectorOptions({ scoreThreshold: detThreshold, inputSize: 320 });
    } catch(e) {
      log('[face-api] ⚠ TinyFaceDetectorOptions constructor error: ' + e.message + ' — pixel fallback.');
      var ctx1 = canvas.getContext('2d');
      var ps1 = { 'top-left': scoreAnchorRegion(ctx1, w, h, 'top-left'), 'top-right': scoreAnchorRegion(ctx1, w, h, 'top-right') };
      ps1['top-left'].textRatio  = detectTextInRegion(ctx1, w, h, 'top-left');
      ps1['top-right'].textRatio = detectTextInRegion(ctx1, w, h, 'top-right');
      return callback(null, { resolvedAnchor: chooseAnchorFromScores(preferredAnchor, ps1), scores: ps1, allUnsafe: false });
    }

    // ── Slides tab: coco-ssd only (faster, better than face-api) ──────────────
    // When coco-ssd is ready and we're not in QYM mode, skip face-api entirely.
    // Fall through to face-api only when coco-ssd is unavailable or for QYM mode.
    if (cocoSsdReady && !_qymFaceRegionOverride) {
      detectPersonsAsync(canvas).then(function (cocoPersons) {
        var textRatioMap = { 'top-left': 0, 'top-right': 0 };
        var anchorKeys0 = ['top-left', 'top-right'];
        if (ocradReady) {
          for (var ak = 0; ak < anchorKeys0.length; ak++) {
            var anchorKey0 = anchorKeys0[ak];
            var reg0 = regionDefs[anchorKey0];
            var hasPerson = false;
            for (var ci = 0; ci < cocoPersons.length; ci++) {
              var cp = cocoPersons[ci];
              if (cp.x2 > reg0.x1 && cp.x1 < reg0.x2 && cp.y2 > reg0.y1 && cp.y1 < reg0.y2) {
                hasPerson = true; break;
              }
            }
            if (!hasPerson) {
              textRatioMap[anchorKey0] = detectTextWithOCRAD(canvas, anchorKey0);
            }
          }
        } else {
          var ctx0 = canvas.getContext('2d');
          textRatioMap['top-left']  = detectTextInRegion(ctx0, canvas.width, canvas.height, 'top-left');
          textRatioMap['top-right'] = detectTextInRegion(ctx0, canvas.width, canvas.height, 'top-right');
        }
        var scores = {};
        anchorKeys0.forEach(function (anchorKey) {
          var reg = regionDefs[anchorKey];
          var faceDetected = false, maxConfidence = 0, nearestFaceDx = Infinity, personCount = 0;
          cocoPersons.forEach(function (p) {
            var gapX = Math.max(reg.x1 - p.x2, 0, p.x1 - reg.x2);
            var gapY = Math.max(reg.y1 - p.y2, 0, p.y1 - reg.y2);
            var dist = Math.sqrt(gapX * gapX + gapY * gapY);
            if (dist < nearestFaceDx) nearestFaceDx = dist;
            if (p.x2 > reg.x1 && p.x1 < reg.x2 && p.y2 > reg.y1 && p.y1 < reg.y2) {
              faceDetected = true; personCount++;
              if (p.score > maxConfidence) maxConfidence = p.score;
            }
          });
          scores[anchorKey] = {
            anchor: anchorKey, faceDetected: faceDetected, maxConfidence: maxConfidence,
            faceCount: personCount, nearestFaceDx: nearestFaceDx,
            textRatio: textRatioMap[anchorKey] || 0,
            skinRatio: 0, score: 0, topSkinRatio: 0, edgeRatio: 0, textureRatio: 0
          };
        });
        callback(null, { resolvedAnchor: chooseAnchorFromScores(preferredAnchor, scores), scores: scores, allUnsafe: false });
      }).catch(function (err) {
        log('[coco-ssd] Slides error: ' + (err && err.message ? err.message : err) + ' — pixel fallback.');
        var ctxE = canvas.getContext('2d');
        var psE = { 'top-left': scoreAnchorRegion(ctxE, w, h, 'top-left'), 'top-right': scoreAnchorRegion(ctxE, w, h, 'top-right') };
        callback(null, { resolvedAnchor: chooseAnchorFromScores(preferredAnchor, psE), scores: psE, allUnsafe: false });
      });
      return;
    }

    // ── face-api path (QYM fallback or when coco-ssd is not loaded) ────────────
    var faceInput = (imgElement && imgElement.naturalWidth) ? imgElement : canvas;
    var iw = (faceInput === imgElement && imgElement.naturalWidth) ? imgElement.naturalWidth  : w;
    var ih = (faceInput === imgElement && imgElement.naturalHeight) ? imgElement.naturalHeight : h;

    faceApiGlobal.detectAllFaces(faceInput, detectorOptions)
      .then(function (detections) {
        var textRatioMap = { 'top-left': 0, 'top-right': 0 };
        if (!_qymFaceRegionOverride) {
          var anchorKeys0 = ['top-left', 'top-right'];
          if (ocradReady) {
            for (var ak = 0; ak < anchorKeys0.length; ak++) {
              var anchorKey0 = anchorKeys0[ak];
              var reg0 = regionDefs[anchorKey0];
              var hasFace = false;
              for (var di = 0; di < detections.length; di++) {
                var det0 = detections[di];
                var bx1 = det0.box.x / iw, by1 = det0.box.y / ih;
                var bx2 = (det0.box.x + det0.box.width)  / iw;
                var by2 = (det0.box.y + det0.box.height) / ih;
                if (bx2 > reg0.x1 && bx1 < reg0.x2 && by2 > reg0.y1 && by1 < reg0.y2) {
                  hasFace = true; break;
                }
              }
              if (!hasFace) {
                textRatioMap[anchorKey0] = detectTextWithOCRAD(canvas, anchorKey0);
              }
            }
          } else {
            var ctx0 = canvas.getContext('2d');
            textRatioMap['top-left']  = detectTextInRegion(ctx0, canvas.width, canvas.height, 'top-left');
            textRatioMap['top-right'] = detectTextInRegion(ctx0, canvas.width, canvas.height, 'top-right');
          }
        }
        return [detections, textRatioMap];
      })
      .then(function (results) {
        var detections   = results[0];
        var textRatioMap = results[1];
        var scores = {};
        var anchorKeys = ['top-left', 'top-right'];

        anchorKeys.forEach(function (anchorKey) {
          var reg = regionDefs[anchorKey];
          var faceDetected = false;
          var maxConfidence = 0;
          var textRatio = textRatioMap[anchorKey] || 0;
          // Distance (in 0–1 frame fractions) from the nearest detected face's
          // bounding box to this anchor's slide region. Used downstream by
          // QYM's classifyQYM to decide whether high TR-skin is "person nearby"
          // (apply tight skin gate) vs "decorative background" (accept).
          var nearestFaceDx = Infinity;

          detections.forEach(function (det) {
            // Normalize detected bounding box to 0–1 fractions of image size
            var bx1 = det.box.x / iw;
            var by1 = det.box.y / ih;
            var bx2 = (det.box.x + det.box.width)  / iw;
            var by2 = (det.box.y + det.box.height) / ih;

            // 2D point-to-rectangle distance: 0 if the boxes overlap.
            var gapX = Math.max(reg.x1 - bx2, 0, bx1 - reg.x2);
            var gapY = Math.max(reg.y1 - by2, 0, by1 - reg.y2);
            var dist = Math.sqrt(gapX * gapX + gapY * gapY);
            if (dist < nearestFaceDx) nearestFaceDx = dist;

            // AABB intersection with the slide placement region
            if (bx2 > reg.x1 && bx1 < reg.x2 && by2 > reg.y1 && by1 < reg.y2) {
              faceDetected = true;
              if (det.score > maxConfidence) maxConfidence = det.score;
            }
          });


          // For QYM mode (large region, may contain heads/hair the face detector
          // misses — profile views, hijabs, faces from behind), additionally run
          // pixel-based skin/edge/texture analysis. This catches heads even when
          // face-api returns nothing.
          var pixelSkin = 0, pixelTopSkin = 0, pixelEdge = 0, pixelTex = 0, pixelScoreVal = 0;
          if (_qymFaceRegionOverride) {
            try {
              var pCtx = canvas.getContext('2d');
              var pAna = scoreAnchorRegion(pCtx, w, h, anchorKey);
              if (pAna) {
                pixelSkin     = pAna.skinRatio    || 0;
                pixelTopSkin  = pAna.topSkinRatio || 0;
                pixelEdge     = pAna.edgeRatio    || 0;
                pixelTex      = pAna.textureRatio || 0;
                pixelScoreVal = pAna.score        || 0;
              }
            } catch (ePix) {}
          }

          // Combine face / text / pixel signals.
          //   faceDetected         → skinRatio 0.9+ (face) — always unsafe
          //   pixelSkin / pixelTex → catches heads / hair the NN missed
          //   text only            → textRatio checked separately
          var faceSkinEquiv = faceDetected ? Math.max(0.9, maxConfidence) : 0.0;
          var skinRatioEquiv = Math.max(faceSkinEquiv, pixelSkin);
          var combinedScore = Math.max(
            (faceDetected || textRatio > 0.15) ? 1.0 : 0.0,
            pixelScoreVal
          );
          scores[anchorKey] = {
            anchor:               anchorKey,
            score:                combinedScore,
            skinRatio:            skinRatioEquiv,
            topSkinRatio:         Math.max(faceSkinEquiv, pixelTopSkin),
            darkRatio:            0,
            lowerOccupiedRatio:   0,
            centralOccupiedRatio: 0,
            clearanceRatio:       0,
            belowBandRatio:       0,
            edgeRatio:            pixelEdge,
            textureRatio:         pixelTex,
            textRatio:            textRatio,
            faceDetected:         faceDetected,
            faceCount:            detections.length,
            maxConfidence:        maxConfidence,
            nearestFaceDx:        nearestFaceDx   // 0–1 frame fraction; Infinity if no faces
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

      // QYM tab: use coco-ssd person detection (whole-body, robust to profile/
      // hijab/back-of-head). Slides tab: keep face-api + pixel pipeline.
      if (_qymFaceRegionOverride && cocoSsdReady) {
        runCocoSsdAnalysis(canvas, preferredAnchor, callback, framePath);
      } else if (faceApiReady) {
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

  // ── QYM debug-frame keeping ───────────────────────────────────────────────
  // When enabled, every analyzed frame is saved to ~/qym-debug/<runStamp>/
  // with the TR region (blue), the inference crop (yellow dashed), and any
  // detected person bboxes (red) drawn on top. Filename encodes the analyzer
  // verdict so you can scan the folder and immediately see what was flagged.
  var _qymDebugDir       = null;   // current run's debug folder, or null when off
  var _qymDebugCounter   = 0;      // monotonically increasing index for filenames

  function _qymInitDebugDir() {
    var keep = !!(qymKeepDebugInput && qymKeepDebugInput.checked);
    if (!keep) { _qymDebugDir = null; return null; }
    try {
      var rootDir = path.join(os.homedir(), 'qym-debug');
      if (!fs.existsSync(rootDir)) fs.mkdirSync(rootDir);
      var ts = new Date();
      var pad = function (n) { return n < 10 ? '0' + n : '' + n; };
      var stamp = ts.getFullYear() + '-' + pad(ts.getMonth() + 1) + '-' + pad(ts.getDate()) +
                  '_' + pad(ts.getHours()) + '-' + pad(ts.getMinutes()) + '-' + pad(ts.getSeconds());
      var runDir = path.join(rootDir, stamp);
      if (!fs.existsSync(runDir)) fs.mkdirSync(runDir);
      _qymDebugDir = runDir;
      _qymDebugCounter = 0;
      return runDir;
    } catch (e) {
      _qymDebugDir = null;
      return null;
    }
  }

  // Draws TR + crop + person boxes on a copy of `srcCanvas` and writes it
  // to disk under _qymDebugDir. Filename format:
  //   NNNN_<framePathBasename>_<VERDICT>[_persons-N][_dx-X.XX].png
  function _qymSaveDebugFrame(srcCanvas, framePath, regionDef, cropRect, people, verdict, personOverlap, nearestDx) {
    if (!_qymDebugDir) return;
    try {
      var w = srcCanvas.width, h = srcCanvas.height;
      var out = document.createElement('canvas');
      out.width = w; out.height = h;
      var ctx = out.getContext('2d');
      ctx.drawImage(srcCanvas, 0, 0);

      // Crop region (yellow dashed) — what coco-ssd actually saw
      ctx.strokeStyle = 'rgba(255,220,0,0.95)';
      ctx.lineWidth = 3;
      ctx.setLineDash([12, 8]);
      ctx.strokeRect(cropRect.x1 * w, cropRect.y1 * h, (cropRect.x2 - cropRect.x1) * w, (cropRect.y2 - cropRect.y1) * h);
      ctx.setLineDash([]);

      // TR region (cyan solid) — where the slide will land
      ctx.strokeStyle = 'rgba(0,200,255,1)';
      ctx.lineWidth = 5;
      ctx.strokeRect(regionDef.x1 * w, regionDef.y1 * h, (regionDef.x2 - regionDef.x1) * w, (regionDef.y2 - regionDef.y1) * h);

      // Detected person bboxes (red = overlapping TR, orange = not)
      for (var i = 0; i < people.length; i++) {
        var p = people[i];
        var overlap = (p.x2 > regionDef.x1 && p.x1 < regionDef.x2 && p.y2 > regionDef.y1 && p.y1 < regionDef.y2);
        ctx.strokeStyle = overlap ? 'rgba(255,40,40,1)' : 'rgba(255,150,40,0.9)';
        ctx.lineWidth = 4;
        ctx.strokeRect(p.x1 * w, p.y1 * h, (p.x2 - p.x1) * w, (p.y2 - p.y1) * h);
        // Confidence label
        ctx.fillStyle = overlap ? 'rgba(255,40,40,1)' : 'rgba(255,150,40,0.9)';
        ctx.font = 'bold 18px Arial';
        ctx.fillText('person ' + (p.score || 0).toFixed(2), p.x1 * w + 4, p.y1 * h + 22);
      }

      // Verdict banner — top-left
      var bannerColor = personOverlap ? 'rgba(220,40,40,0.92)' : 'rgba(40,170,60,0.92)';
      ctx.fillStyle = bannerColor;
      ctx.fillRect(20, 20, 720, 64);
      ctx.fillStyle = 'white';
      ctx.font = 'bold 26px Arial';
      ctx.fillText(verdict, 36, 60);
      ctx.font = '16px Arial';
      ctx.fillText('persons=' + people.length + '  dx=' + (isFinite(nearestDx) ? nearestDx.toFixed(2) : '∞'), 36, 80);

      _qymDebugCounter++;
      var idxStr = ('0000' + _qymDebugCounter).slice(-4);
      var baseName = framePath ? path.basename(framePath, path.extname(framePath)) : 'frame';
      var verdictTag = personOverlap ? 'REJECT' : 'SAFE';
      var fileName = idxStr + '_' + baseName + '_' + verdictTag +
                     (people.length ? '_p' + people.length : '') +
                     (isFinite(nearestDx) ? '_dx' + nearestDx.toFixed(2) : '') + '.png';
      var fullPath = path.join(_qymDebugDir, fileName);
      var dataURL = out.toDataURL('image/png');
      var b64 = dataURL.replace(/^data:image\/png;base64,/, '');
      fs.writeFileSync(fullPath, b64, 'base64');
    } catch (e) { /* swallow — never let debug saves break a run */ }
  }

  // ── QYM person-detection analyzer (coco-ssd, CROPPED) ─────────────────────
  // Runs SSD MobileNet on a CROP of the frame around the slide region (TR)
  // with a 10% margin of body context. Then checks whether any detected
  // person's bbox (mapped back to full-frame coords) intersects the actual
  // slide region with a 3% safety expansion.
  //
  // Why crop instead of full-frame inference:
  //   - Full-frame coco-ssd produces loose bboxes that bleed into TR even
  //     when the person is centered far from the slide region. The model
  //     then wrongly flags "person in TR" → false rejections.
  //   - On a crop containing only the slide region (+ margin), the model
  //     can only detect persons whose bodies are actually in or right at
  //     the edge of TR.
  //   - Bonus: ~5–6% of original frame area → ~10× faster inference.
  //
  // The 10% margin gives the model enough body context (it was trained on
  // full-body images and may miss isolated heads/arms in tiny crops).
  var QYM_CROP_MARGIN     = 0.10;   // 10% margin of body context around TR
  var QYM_OVERLAP_EXPAND  = 0.03;   // 3% safety buffer on TR overlap check

  function runCocoSsdAnalysis(canvas, preferredAnchor, callback, framePath) {
    var w = canvas.width, h = canvas.height;
    var regionDefs = {
      'top-left':  { x1: 0.15, y1: 0.04, x2: 0.43, y2: 0.13 },
      'top-right': { x1: 0.66, y1: 0.02, x2: 1.00, y2: 0.09 }
    };
    if (_qymFaceRegionOverride) {
      if (_qymFaceRegionOverride['top-left'])  regionDefs['top-left']  = _qymFaceRegionOverride['top-left'];
      if (_qymFaceRegionOverride['top-right']) regionDefs['top-right'] = _qymFaceRegionOverride['top-right'];
    }

    // Only the preferredAnchor (top-right for QYM) gets a real inference —
    // QYM doesn't use top-left, so we save a second model run.
    var reg = regionDefs[preferredAnchor];
    if (!reg) {
      // Shouldn't happen for QYM (always top-right), but defend anyway.
      callback(null, {
        resolvedAnchor: preferredAnchor,
        reason:         'coco-ssd-no-region',
        scores:         { 'top-left': _emptyCocoScore('top-left'), 'top-right': _emptyCocoScore('top-right') },
        allUnsafe:      false
      });
      return;
    }

    // Compute crop window (TR + margin), clamped to frame
    var crX1 = Math.max(0, reg.x1 - QYM_CROP_MARGIN);
    var crY1 = Math.max(0, reg.y1 - QYM_CROP_MARGIN);
    var crX2 = Math.min(1, reg.x2 + QYM_CROP_MARGIN);
    var crY2 = Math.min(1, reg.y2 + QYM_CROP_MARGIN);
    var cropW = Math.round((crX2 - crX1) * w);
    var cropH = Math.round((crY2 - crY1) * h);

    if (cropW <= 1 || cropH <= 1) {
      // Degenerate crop — fall through with empty scores.
      callback(null, {
        resolvedAnchor: preferredAnchor,
        reason:         'coco-ssd-crop-empty',
        scores:         { 'top-left': _emptyCocoScore('top-left'), 'top-right': _emptyCocoScore('top-right') },
        allUnsafe:      false
      });
      return;
    }

    // Build the crop canvas
    var cropCanvas = document.createElement('canvas');
    cropCanvas.width  = cropW;
    cropCanvas.height = cropH;
    cropCanvas.getContext('2d').drawImage(
      canvas,
      crX1 * w, crY1 * h, (crX2 - crX1) * w, (crY2 - crY1) * h,   // source rect
      0, 0, cropW, cropH                                            // dest rect
    );

    detectPersonsAsync(cropCanvas).then(function (peopleInCrop) {
      // Map detections from CROP coords (0–1 of crop) back to FULL-FRAME
      // coords (0–1 of original frame), so overlap checks reuse the same
      // regionDef the rest of the code thinks in.
      var people = peopleInCrop.map(function (p) {
        return {
          x1: crX1 + p.x1 * (crX2 - crX1),
          y1: crY1 + p.y1 * (crY2 - crY1),
          x2: crX1 + p.x2 * (crX2 - crX1),
          y2: crY1 + p.y2 * (crY2 - crY1),
          score: p.score
        };
      });

      // Overlap check against the ORIGINAL TR (with 3% expansion).
      var ex1 = reg.x1 - QYM_OVERLAP_EXPAND, ey1 = reg.y1 - QYM_OVERLAP_EXPAND;
      var ex2 = reg.x2 + QYM_OVERLAP_EXPAND, ey2 = reg.y2 + QYM_OVERLAP_EXPAND;

      var personOverlap     = false;
      var overlapConfidence = 0;
      var nearestDx         = Infinity;

      for (var i = 0; i < people.length; i++) {
        var p = people[i];
        if (p.x2 > ex1 && p.x1 < ex2 && p.y2 > ey1 && p.y1 < ey2) {
          personOverlap = true;
          if (p.score > overlapConfidence) overlapConfidence = p.score;
        }
        var gapX = Math.max(reg.x1 - p.x2, 0, p.x1 - reg.x2);
        var gapY = Math.max(reg.y1 - p.y2, 0, p.y1 - reg.y2);
        var dist = Math.sqrt(gapX * gapX + gapY * gapY);
        if (dist < nearestDx) nearestDx = dist;
      }

      // ── Graphic overlay check ──────────────────────────────────────────
      // Only run when coco-ssd already says "no person in TR".
      // Uses hard-edge density (not OCR) — works for any graphic shape/color.
      var textDetected = false;
      var recognizedText = '';
      if (!personOverlap) {
        var ocrResult = _qymDetectGraphicInRegion(canvas);
        if (ocrResult && ocrResult.detected) {
          textDetected = true;
          recognizedText = ocrResult.text;
        }
      }

      var scores = {
        'top-left':  _emptyCocoScore('top-left'),
        'top-right': _emptyCocoScore('top-right')
      };
      scores[preferredAnchor] = {
        anchor:        preferredAnchor,
        faceDetected:  personOverlap,
        maxConfidence: overlapConfidence,
        faceCount:     people.length,
        nearestFaceDx: nearestDx,
        textDetected:  textDetected,
        recognizedText: recognizedText,
        skinRatio:     0,
        score:         0,
        textRatio:     textDetected ? 0.9 : 0,
        topSkinRatio:  0,
        edgeRatio:     0,
        textureRatio:  0
      };

      // Optional debug-frame save with annotations (TR + crop + person bboxes)
      if (_qymDebugDir) {
        var verdict = personOverlap
          ? 'PERSON-IN-TR conf=' + overlapConfidence.toFixed(2)
          : (textDetected
              ? 'TEXT-IN-TR "' + recognizedText.slice(0, 40) + '"'
              : 'TR-CLEAR');
        _qymSaveDebugFrame(
          canvas, framePath, reg,
          { x1: crX1, y1: crY1, x2: crX2, y2: crY2 },
          people, verdict, personOverlap || textDetected, nearestDx
        );
      }

      callback(null, {
        resolvedAnchor: preferredAnchor,
        reason:         'coco-ssd-cropped',
        scores:         scores,
        allUnsafe:      false,
        faceCount:      people.length
      });
    });
  }

  function _emptyCocoScore(anchorKey) {
    return {
      anchor: anchorKey,
      faceDetected: false,
      maxConfidence: 0,
      faceCount: 0,
      nearestFaceDx: Infinity,
      textDetected: false,
      recognizedText: '',
      skinRatio: 0, score: 0, textRatio: 0,
      topSkinRatio: 0, edgeRatio: 0, textureRatio: 0
    };
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
        faceCount: 0,
        nearestFaceDx: Infinity,
        textDetected: false,
        recognizedText: '',
        unsafeSampleCount: 0,
        totalSampleCount: 0
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
        faceCount: 0,
        nearestFaceDx: Infinity,
        unsafeSampleCount: 0,
        totalSampleCount: 0
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
        // Per-sample tally so we can apply majority voting (QYM) instead of MAX (Slides tab default).
        target.totalSampleCount++;
        if (source.faceDetected) target.unsafeSampleCount++;
        if ((source.maxConfidence || 0) > target.maxConfidence) target.maxConfidence = source.maxConfidence;
        if ((source.faceCount || 0) > target.faceCount) target.faceCount = source.faceCount;
        // Worst case across frames = closest face seen in any sample
        if (typeof source.nearestFaceDx === 'number' && source.nearestFaceDx < target.nearestFaceDx) {
          target.nearestFaceDx = source.nearestFaceDx;
        }
        // Text detection (QYM) — OR across samples; remember any recognised text for the log.
        if (source.textDetected) {
          target.textDetected = true;
          if (source.recognizedText && !target.recognizedText) {
            target.recognizedText = source.recognizedText;
          }
        }
      });
    });

    // Decide final faceDetected:
    //   ANY sample unsafe → reject the whole candidate.
    //
    // Tried majority voting (2/3 unsafe to reject), but that accepted candidates
    // where 1 of 3 sampled moments had a person — meaning ~3 seconds of slide
    // time visibly sits over a head/face. Strict is the right behaviour here:
    // an opaque ad card briefly covering then revealing a person looks bad
    // even if it's only ~33% of the display window.
    //
    // The samples=N/M counter is kept in the trace log for diagnostics.
    ['top-left', 'top-right'].forEach(function (anchorKey) {
      var t = merged[anchorKey];
      t.faceDetected = (t.unsafeSampleCount > 0);
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
    // Minimum gap between consecutive slides. In pack mode the user wants a
    // visible breathing space between back-to-back slides; in spread mode the
    // existing 0.1s prevents only exact overlap.
    var slideGap = options.packPerCategory ? 0.5 : 0.1;
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

    // ── Sort placements within each category by startSeconds ────────────────
    // English is always files[0] in the batch; sorting ensures it maps to the
    // earliest timecode in the category regardless of gap-placement reordering.
    function sortCategoryPlacementsByTime() {
      var groups = {}, order = [];
      placements.forEach(function (p) {
        var c = p.categoryName || '_';
        if (!groups[c]) { groups[c] = []; order.push(c); }
        groups[c].push(p);
      });
      var sorted = [];
      order.forEach(function (c) {
        groups[c].sort(function (a, b) { return a.startSeconds - b.startSeconds; });
        groups[c].forEach(function (p) { sorted.push(p); });
      });
      placements.length = 0;
      sorted.forEach(function (p) { placements.push(p); });
    }

    // ── Main placement loop ──────────────────────────────────────────────────
    var prevCatName = null; // track category changes to trigger deferred retry

    function placeNext() {
      if (index >= placements.length) {
        // Retry any deferred slides for the last category before finishing
        if (prevCatName && deferredSlides[prevCatName] && deferredSlides[prevCatName].length) {
          retryDeferred(prevCatName, function () {
            sortCategoryPlacementsByTime();
            callback(null, placementPlan, { movedCount: movedCount, unsafeFallbackCount: unsafeFallbackCount });
          });
        } else {
          sortCategoryPlacementsByTime();
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
      } else if (options.packPerCategory) {
        // Pack mode: each slide hugs the previous one (cursor already includes
        // slideGap from the prior placement). Inter-group gap comes for free
        // from the zone-bounded cursor reset on the next category's first slide.
        originalStart = cursor;
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

  var _runBtnDefaultLabel = null;
  function setRunning(on) {
    if (_runBtnDefaultLabel === null) _runBtnDefaultLabel = runBtn.innerHTML;
    runBtn.disabled  = !!on;
    runBtn.innerHTML = on ? 'Working…' : _runBtnDefaultLabel;
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

      // Pack mode: tight in-row placement takes priority over V1 avoidance
      // and face/head avoidance. Clear blocked V1 ranges so the placement
      // formula isn't nudged off its computed positions, and force
      // avoidFaces=false so the OFF (formula) path runs instead of the
      // frame-by-frame face/head search.
      if (options.packPerCategory) {
        if (previewParsed.placementPlan) {
          previewParsed.placementPlan.blockedV1Ranges = [];
        }
        if (options.avoidFaces) {
          options.avoidFaces = false;
          log('Pack mode: face/head detection skipped — slides land at packed positions.');
        }
        log('Pack mode: V1 avoidance disabled — slides may overlap V1 clips.');
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

      // No face avoidance — redistribute slides within each category zone with a
      // uniform step (zone.length / n), starting the first slide at zone.start.
      // The raw JSX formula uses (p+1)*interval so slide 0 starts one interval in
      // and the gap from slide 0 → slide 1 ends up double-width after anchoring
      // only slide 0.  Recomputing all positions with step j*(zone/n) gives equal
      // intervals across the whole category.
      // V1 blocked ranges are nudged for each slide individually.
      if (previewParsed.placementPlan) {
        var plan0 = previewParsed.placementPlan;
        var pmts0 = plan0.placements || [];
        var blk0  = Array.isArray(plan0.blockedV1Ranges) ? plan0.blockedV1Ranges : [];
        var wSt0  = typeof plan0.placementWindowStartSeconds === 'number' ? plan0.placementWindowStartSeconds : 0;
        var wEd0  = (typeof plan0.placementWindowEndSeconds  === 'number' && plan0.placementWindowEndSeconds  > 0)
                    ? plan0.placementWindowEndSeconds
                    : (typeof plan0.usedTimelineLengthSeconds === 'number' ? plan0.usedTimelineLengthSeconds : 0);

        if (wEd0 > wSt0 && pmts0.length > 0) {
          // Build category list in order of first appearance
          var catNames0 = [];
          var seenC0    = {};
          pmts0.forEach(function (p) {
            var c = p.categoryName || '_';
            if (!seenC0[c]) { seenC0[c] = true; catNames0.push(c); }
          });
          var usable0  = plan0.usableTimelineLengthSeconds || (wEd0 - wSt0);
          var perCat0  = usable0 / (catNames0.length || 1);

          // Walk forward skipping V1 blocks — same as mapCompressedToReal in buildSafePlacementPlan
          function mapC0(csec) {
            if (!blk0.length) return Math.min(wSt0 + csec, wEd0);
            var srt = blk0.slice().sort(function (a, b) { return a.start - b.start; });
            var rem = csec, t = wSt0;
            for (var r = 0; r < srt.length; r++) {
              var bs = Math.max(srt[r].start, wSt0), be = Math.min(srt[r].end, wEd0);
              if (bs >= wEd0) break;
              var fr = Math.max(0, bs - t);
              if (rem <= fr) return t + rem;
              rem -= fr; t = Math.max(t, be);
            }
            return Math.min(t + rem, wEd0);
          }

          // Per-category zones (same boundaries as buildSafePlacementPlan)
          var catZones0 = {};
          catNames0.forEach(function (c, i) {
            catZones0[c] = { start: mapC0(i * perCat0), end: mapC0((i + 1) * perCat0) };
          });

          // Slide count per category
          var catTotal0 = {};
          pmts0.forEach(function (p) {
            var c = p.categoryName || '_';
            catTotal0[c] = (catTotal0[c] || 0) + 1;
          });

          // Running index per category (used as j in the step formula)
          var catIdx0 = {};
          catNames0.forEach(function (c) { catIdx0[c] = 0; });

          // Apply positions per category. Default: equal-step
          //   slide j → zone.start + j * (zone.length / n)
          // Pack mode: tight back-to-back from zone.start, with a per-category
          // cursor so V1-blocked nudges accumulate (otherwise consecutive slides
          // can collapse onto the same nextAvailableStart and overlap by ~9s).
          var packMode0 = !!options.packPerCategory;
          var slideGap0 = packMode0 ? 0.5 : 0.1;
          var catCursor0 = {};
          catNames0.forEach(function (c) {
            catCursor0[c] = (catZones0[c] || { start: wSt0 }).start;
          });
          pmts0.forEach(function (p, pi) {
            var cat  = p.categoryName || '_';
            var zone = catZones0[cat] || { start: wSt0, end: wEd0 };
            var n    = catTotal0[cat] || 1;
            var j    = catIdx0[cat];
            var dur  = p.clipDurationSeconds || 9;
            var t;
            if (packMode0) {
              // Cursor was initialised to zone.start, then advanced by
              // (dur + slideGap0) after each placement — guarantees no overlap
              // even when earlier slides were nudged out of V1 ranges.
              t = catCursor0[cat];
            } else {
              var step = (zone.end - zone.start) / n;
              t = zone.start + j * step;
            }
            if (blk0.length && overlapsBlockedRange(t, dur, blk0)) {
              t = nextAvailableStart(t, dur, blk0);
            }
            if (Math.abs(p.startSeconds - t) > 0.1) {
              log('Slide ' + (pi + 1) + ' [' + cat + ']: ' +
                  (packMode0 ? 'packed' : 'equally spaced') + ' at ' + secToMS(t) + '.');
            }
            p.startSeconds = t;
            catIdx0[cat]++;
            if (packMode0) catCursor0[cat] = t + dur + slideGap0;
          });
        }
      }

      placeResolvedPlan(previewParsed.placementPlan, { movedCount: 0, unsafeFallbackCount: 0 });
    });
  }

  // ════════════════════════════════════════════════════════════════════════
  //  QUAN-YIN & SMTV MAX TAB
  // ════════════════════════════════════════════════════════════════════════

  // ── Tab switching ───────────────────────────────────────────────────────
  function switchTab(name) {
    var panelSlides = document.getElementById('panelSlides');
    var panelQYM    = document.getElementById('panelQYM');
    var btnSlides   = document.getElementById('tabSlidesBtn');
    var btnQYM      = document.getElementById('tabQYMBtn');
    if (panelSlides) panelSlides.classList.toggle('active', name === 'Slides');
    if (panelQYM)    panelQYM.classList.toggle('active',    name === 'QYM');
    if (btnSlides)   btnSlides.classList.toggle('active',   name === 'Slides');
    if (btnQYM)      btnQYM.classList.toggle('active',      name === 'QYM');
  }
  document.getElementById('tabSlidesBtn').addEventListener('click', function () { switchTab('Slides'); });
  document.getElementById('tabQYMBtn').addEventListener('click',    function () { switchTab('QYM'); });

  // ── QYM helpers ─────────────────────────────────────────────────────────
  function logQYM(msg) {
    if (!qymStatusEl) return;
    qymStatusEl.textContent += '\n' + msg;
    qymStatusEl.scrollTop = qymStatusEl.scrollHeight;
  }
  function setQymStatus(msg) {
    if (qymStatusEl) qymStatusEl.textContent = msg;
  }
  // Original button label so we can restore it when the run finishes.
  var qymRunBtnDefaultLabel = null;
  function setQymRunning(running) {
    qymRunning = running;
    if (qymRunBtn) {
      if (qymRunBtnDefaultLabel === null) qymRunBtnDefaultLabel = qymRunBtn.innerHTML;
      qymRunBtn.disabled  = running;
      qymRunBtn.innerHTML = running ? 'Working…' : qymRunBtnDefaultLabel;
    }
  }

  // ── Settings persist / restore ──────────────────────────────────────────
  function persistQYMSettings() {
    var tracking = loadTracking();
    tracking.settings.qymQuanYinFolder = selectedQYFolder  || '';
    tracking.settings.qymSmtvMaxFolder = selectedMaxFolder || '';
    tracking.settings.qymMaxCount      = parseInt(qymMaxCountInput.value, 10)      || 4;
    tracking.settings.qymTargetTrack   = parseInt(qymTargetTrackInput.value, 10)   || 10;
    tracking.settings.qymFadeDuration  = parseFloat(qymFadeDurationInput.value)    || 0.3;
    tracking.settings.qymAvoidFaces    = !!(qymAvoidFacesInput && qymAvoidFacesInput.checked);
    tracking.settings.qymSlideType     = getQymSlideType();
    tracking.settings.qymKeepDebug     = !!(qymKeepDebugInput && qymKeepDebugInput.checked);
    saveTracking(tracking);
  }
  function restoreQYMSettings() {
    var tracking = loadTracking();
    if (tracking.settings.qymQuanYinFolder) {
      selectedQYFolder = tracking.settings.qymQuanYinFolder;
      if (qymQYFolderInput) qymQYFolderInput.value = selectedQYFolder;
    }
    if (tracking.settings.qymSmtvMaxFolder) {
      selectedMaxFolder = tracking.settings.qymSmtvMaxFolder;
      if (qymMaxFolderInput) qymMaxFolderInput.value = selectedMaxFolder;
    }
    if (qymMaxCountInput)     qymMaxCountInput.value     = tracking.settings.qymMaxCount     || 4;
    if (qymTargetTrackInput)  qymTargetTrackInput.value  = tracking.settings.qymTargetTrack  || 10;
    if (qymFadeDurationInput) qymFadeDurationInput.value = tracking.settings.qymFadeDuration !== undefined ? tracking.settings.qymFadeDuration : 0.3;
    if (qymAvoidFacesInput)   qymAvoidFacesInput.checked = tracking.settings.qymAvoidFaces !== false;
    if (qymKeepDebugInput)    qymKeepDebugInput.checked  = !!tracking.settings.qymKeepDebug;
    setQymSlideType(tracking.settings.qymSlideType || 'both');
  }

  // ── Folder scanning ─────────────────────────────────────────────────────

  function scanQuanYinFolder(folderPath) {
    // Returns { englishA, englishB, nonEnglish: [{lang, filePath}] }
    var result = { englishA: null, englishB: null, nonEnglish: [] };
    if (!folderPath || !fs.existsSync(folderPath)) return result;
    var files;
    try { files = fs.readdirSync(folderPath); } catch (e) { return result; }
    files.forEach(function (name) {
      if (path.extname(name).toLowerCase() !== '.png') return;
      var full = path.join(folderPath, name);
      var upper = name.toUpperCase();
      if (upper.indexOf('ENGLISH A') !== -1) { result.englishA = full; return; }
      if (upper.indexOf('ENGLISH B') !== -1) { result.englishB = full; return; }
      // Extract language from everything after '--'
      var sep = name.indexOf('--');
      if (sep === -1) return; // skip unrecognised files
      var lang = name.slice(sep + 2).replace(/\.png$/i, '').trim();
      if (lang) result.nonEnglish.push({ lang: lang, filePath: full });
    });
    return result;
  }

  function scanSmtvMaxFolder(folderPath) {
    // Returns { english, nonEnglish: [{lang, filePath}] }
    var result = { english: null, nonEnglish: [] };
    if (!folderPath || !fs.existsSync(folderPath)) return result;
    var files;
    try { files = fs.readdirSync(folderPath); } catch (e) { return result; }
    files.forEach(function (name) {
      if (path.extname(name).toLowerCase() !== '.png') return;
      var full   = path.join(folderPath, name);
      var upper  = name.toUpperCase();
      // English: filename contains '-ENG.' (not TELU, HINDI etc.)
      if (upper.match(/-ENG\./)) { result.english = full; return; }
      // Language code: extract after last '-', drop extension
      var lastDash = name.lastIndexOf('-');
      if (lastDash === -1) return;
      var lang = name.slice(lastDash + 1).replace(/\.png$/i, '').trim();
      if (lang) result.nonEnglish.push({ lang: lang, filePath: full });
    });
    return result;
  }

  // ── Language cycle picking ──────────────────────────────────────────────

  // Normalise a language label (Quan-Yin uses full names like "Persian",
  // SMTV Max uses short codes like "FAR") to a canonical key so we can detect
  // overlap between the two folders.
  // Verified language alias map — every entry below corresponds to an actual
  // file in either W:\...\Quan Yin or W:\...\SMTV MAX. The QY full name and
  // its SMTV Max code map to the same canonical key so the overlap-avoidance
  // check in pickCycleNonEnglish can detect equivalence.
  //
  //   QY filename: "Quan-Yin？--<Name>.png"     SMTV Max: "SMTV Max-<CODE>.png"
  //   Both forms get lowercased + non-alpha stripped before lookup.
  //
  // Pairs verified against folder listing (29 QY langs, 38 SMTV codes):
  //   Arabic↔ARA  Aulacese↔AUL  Bulgarian↔BUL  Chinese Trad↔CHI TRAD
  //   Czech↔CZE   French↔FRE    German↔GER     Hindi↔HINDI
  //   Hungarian↔HUN  Indonesian↔INDO  Italian↔ITA  Japanese↔JAP
  //   Korean↔KOR  Malay↔MALAY   Mongolian↔MON  Persian↔PER
  //   Polish↔POL  Portuguese↔POR  Punjabi↔PUN  Romanian↔ROM
  //   Russian↔RUS  Spanish↔SPA  Telugu↔TELU    Thai↔THAI
  //   Ukrainian↔UKR  Urdu↔URD
  //   QY-only: Chinese Simp
  //   SMTV-only: CRO DUT EST EWE FIN NEP NOR SLO SWE TUR ZUL ENG
  var _QYM_LANG_ALIAS = {
    // ── SMTV Max short codes ──────────────────────────────────────────
    'ara':'ar',     'aul':'vi',     'bul':'bg',     'chitrad':'zh-trad',
    'cro':'hr',     'cze':'cs',     'dut':'nl',     'eng':'en',
    'est':'et',     'ewe':'ee',     'fin':'fi',     'fre':'fr',
    'ger':'de',     'hindi':'hi',   'hun':'hu',     'indo':'id',
    'ita':'it',     'jap':'ja',     'kor':'ko',     'malay':'ms',
    'mon':'mn',     'nep':'ne',     'nor':'no',     'per':'fa',
    'pol':'pl',     'por':'pt',     'pun':'pa',     'rom':'ro',
    'rus':'ru',     'slo':'sk',     'spa':'es',     'swe':'sv',
    'telu':'te',    'thai':'th',    'tur':'tr',     'ukr':'uk',
    'urd':'ur',     'zul':'zu',
    // ── Quan-Yin full names (lowercase, non-alpha stripped) ──
    'arabic':'ar',       'aulacese':'vi',     'bulgarian':'bg',
    'chinesesimp':'zh',  'chinesetrad':'zh-trad',
    'czech':'cs',        'french':'fr',       'german':'de',
    'hindi2':'hi', /* unreachable — 'hindi' matches first */
    'hungarian':'hu',    'indonesian':'id',   'italian':'it',
    'japanese':'ja',     'korean':'ko',       'malay2':'ms', /* unreachable */
    'mongolian':'mn',    'persian':'fa',      'polish':'pl',
    'portuguese':'pt',   'punjabi':'pa',      'romanian':'ro',
    'russian':'ru',      'spanish':'es',      'telugu':'te',
    'thai2':'th', /* unreachable */
    'ukrainian':'uk',    'urdu':'ur',
    // ── Common synonyms (defensive, in case a future filename uses them) ──
    'farsi':'fa',        'vietnamese':'vi',   'english':'en',
    'dutch':'nl',        'holland':'nl',      'norwegian':'no',
    'swedish':'sv',      'finnish':'fi',      'estonian':'et',
    'croatian':'hr',     'slovak':'sk',       'slovenian':'sl',
    'turkish':'tr',      'zulu':'zu',         'nepali':'ne',
    'chinese':'zh'
  };
  function _qymCanonLang(name) {
    if (!name) return '';
    var s = String(name).toLowerCase().replace(/[^a-z]/g, '');
    if (!s) return '';
    if (_QYM_LANG_ALIAS[s]) return _QYM_LANG_ALIAS[s];
    // Last-resort fallback for unknown labels: 3-letter prefix.
    return s.slice(0, 3);
  }

  // pickCycleNonEnglish picks `needed` languages while respecting:
  //   1. The cycle history in usedCycle (avoids repeats across runs)
  //   2. excludeCanonSet (avoids overlap with the OTHER type in this run)
  //
  // Tier order:
  //   Tier 1 — unused-this-cycle  AND  not-overlapping-with-other-type
  //   Tier 2 — restart cycle but stay non-overlapping (cycle exhausted, still avoid overlap)
  //   Tier 3 — fallback: allow overlap (only when avoidance is impossible)
  function pickCycleNonEnglish(allNonEnglish, needed, usedCycle, labelForLog, excludeCanonSet) {
    if (needed <= 0) return [];
    var used = Array.isArray(usedCycle) ? usedCycle : [];
    var hasExclude = excludeCanonSet && typeof excludeCanonSet === 'object';
    function isExcluded(item) {
      if (!hasExclude) return false;
      return excludeCanonSet[_qymCanonLang(item.lang)] === true;
    }

    // Tier 1: not-in-cycle AND not-excluded
    var tier1 = allNonEnglish.filter(function (it) {
      return used.indexOf(it.lang) === -1 && !isExcluded(it);
    });
    if (tier1.length >= needed) {
      var picked1 = shuffle(tier1).slice(0, needed);
      picked1.forEach(function (it) { if (used.indexOf(it.lang) === -1) used.push(it.lang); });
      return picked1;
    }

    // Tier 2: cycle exhausted but we can still avoid overlap — restart cycle
    var nonExcludedAll = allNonEnglish.filter(function (it) { return !isExcluded(it); });
    if (nonExcludedAll.length >= needed) {
      logQYM(labelForLog + ': non-English cycle complete — restarting (still avoiding overlap).');
      used.splice(0, used.length);
      var picked2 = shuffle(nonExcludedAll).slice(0, needed);
      picked2.forEach(function (it) { if (used.indexOf(it.lang) === -1) used.push(it.lang); });
      return picked2;
    }

    // Tier 3: cannot fully avoid overlap — fill non-overlapping first, then overlap
    var picked3 = nonExcludedAll.slice();
    var taken = {};
    picked3.forEach(function (it) { taken[it.filePath] = true; });
    var remaining = allNonEnglish.filter(function (it) { return !taken[it.filePath]; });
    var moreNeeded = needed - picked3.length;
    if (moreNeeded > 0) {
      picked3 = picked3.concat(shuffle(remaining).slice(0, moreNeeded));
      if (hasExclude) {
        logQYM(labelForLog + ': only ' + nonExcludedAll.length + ' non-overlapping language(s) available — allowing overlap with Quan-Yin for the remaining ' + moreNeeded + '.');
      } else {
        logQYM(labelForLog + ': non-English cycle complete — restarting.');
      }
    }
    used.splice(0, used.length);
    picked3.forEach(function (it) { if (used.indexOf(it.lang) === -1) used.push(it.lang); });
    return picked3.slice(0, needed);
  }

  function pickQuanYinSlides(scan, maxCount, tracking) {
    var warnings = [];
    // English (A/B toggle)
    var next = (tracking.quanYinCycle && tracking.quanYinCycle.nextEnglish) || 'A';
    var englishFile = (next === 'A') ? scan.englishA : scan.englishB;
    if (!englishFile) {
      englishFile = scan.englishA || scan.englishB;
      warnings.push('Quan-Yin ENGLISH ' + next + ' file not found — using available English variant.');
    }
    // Toggle for next run
    if (!tracking.quanYinCycle) tracking.quanYinCycle = { nextEnglish: 'A', usedNonEnglish: [] };
    tracking.quanYinCycle.nextEnglish = (next === 'A') ? 'B' : 'A';

    var needed = Math.max(0, maxCount - 1);
    var picked = pickCycleNonEnglish(scan.nonEnglish, needed, tracking.quanYinCycle.usedNonEnglish, 'Quan-Yin');

    var files = [];
    var isEnglishArr = [];
    var langs = [];
    if (englishFile) { files.push(englishFile); isEnglishArr.push(true);  langs.push('English'); }
    picked.forEach(function (item) { files.push(item.filePath); isEnglishArr.push(false); langs.push(item.lang); });

    if (files.length === 0) warnings.push('Quan-Yin: no files could be selected.');
    return { files: files, isEnglish: isEnglishArr, langs: langs, warnings: warnings };
  }

  function pickSmtvMaxSlides(scan, maxCount, tracking, excludeCanonSet) {
    var warnings = [];
    if (!scan.english) warnings.push('SMTV Max: English file not found.');

    var needed = Math.max(0, maxCount - 1);
    if (!tracking.smtvMaxCycle) tracking.smtvMaxCycle = { usedNonEnglish: [] };
    var picked = pickCycleNonEnglish(scan.nonEnglish, needed, tracking.smtvMaxCycle.usedNonEnglish, 'SMTV Max', excludeCanonSet);

    var files = [];
    var isEnglishArr = [];
    var langs = [];
    if (scan.english) { files.push(scan.english); isEnglishArr.push(true);  langs.push('English'); }
    picked.forEach(function (item) { files.push(item.filePath); isEnglishArr.push(false); langs.push(item.lang); });

    if (files.length === 0) warnings.push('SMTV Max: no files could be selected.');
    return { files: files, isEnglish: isEnglishArr, langs: langs, warnings: warnings };
  }

  // ── Interleave QY + Max slides ──────────────────────────────────────────
  // Returns flat array: [QY-0, MAX-0, QY-1, MAX-1, ...]
  function interleaveQYM(qyResult, maxResult) {
    var out = [];
    var qn = qyResult.files.length;
    var mn = maxResult.files.length;
    var len = Math.max(qn, mn);
    for (var i = 0; i < len; i++) {
      if (i < qn) out.push({ filePath: qyResult.files[i],  type: 'quan-yin',  isEnglish: qyResult.isEnglish[i],  lang: qyResult.langs[i] });
      if (i < mn) out.push({ filePath: maxResult.files[i], type: 'smtv-max', isEnglish: maxResult.isEnglish[i], lang: maxResult.langs[i] });
    }
    return out;
  }

  // ── Equal-step positions (avoidFaces = OFF) ──────────────────────────────
  function computeQYMPositions(slides, windowStart, windowEnd, clipDuration) {
    var n    = slides.length;
    var step = n > 1 ? (windowEnd - windowStart) / n : 0;
    return slides.map(function (_, j) {
      return { startSeconds: windowStart + j * step };
    });
  }

  // ── Safe placement plan for QYM (avoidFaces = ON) ────────────────────────
  // Single pool, no category zones — same candidate search as buildSafePlacementPlan
  // but simplified: fixed corner top-right, no zone limits, no deferred slides.
  function buildSafePlacementPlanQYM(options, slides, windowStart, windowEnd, callback) {
    if (!slides.length) { callback(null, []); return; }
    // clipDuration may be shortened (for sub-9s windows) — see runQuanYinAndMax.
    var customClipDuration = (options && typeof options.clipDuration === 'number') ? options.clipDuration : null;

    // ── Override the analyzer's tiny "top badge" region with the actual ─────
    // QYM ad-card footprint. JSX presets place both Quan-Yin and SMTV Max at
    // position [1688, 440] on a 1920×1080 sequence, scale 90%:
    //   Quan-Yin  ~488×251 src → displayed 439×226 → bounds x=[0.765,0.993], y=[0.303,0.512]
    //   SMTV Max  ~446×273 src → displayed 401×246 → bounds x=[0.775,0.984], y=[0.293,0.521]
    // Use the union (a tiny safety margin on each side) so face detection looks
    // at the area where the slide actually lands, not the empty top strip.
    _qymFaceRegionOverride = {
      'top-right': { x1: 0.760, y1: 0.288, x2: 0.998, y2: 0.526 },
      'top-left':  { x1: 0.002, y1: 0.288, x2: 0.240, y2: 0.526 } // mirror, unused for QYM
    };
    logQYM('  [face-region] using QYM footprint: x=' +
           (_qymFaceRegionOverride['top-right'].x1 * 100).toFixed(0) + '–' +
           (_qymFaceRegionOverride['top-right'].x2 * 100).toFixed(0) + '%, y=' +
           (_qymFaceRegionOverride['top-right'].y1 * 100).toFixed(0) + '–' +
           (_qymFaceRegionOverride['top-right'].y2 * 100).toFixed(0) + '%');
    logQYM('  [detector] ' + (cocoSsdReady
      ? 'coco-ssd PERSON (TR crop ±10%) + ' + (ocradReady
          ? 'OCRAD TEXT (TR + 0/0/+10/+10 margins) when person-clear'
          : 'OCRAD unavailable (text check skipped)')
      : (faceApiReady
          ? 'face-api FACE detection only (coco-ssd not loaded)'
          : 'pixel-based fallback (no NN models loaded)')));

    // Wrap the user's callback so the override is always cleared on exit.
    var _origCallback = callback;
    callback = function (err, res) {
      _qymFaceRegionOverride = null;
      _origCallback(err, res);
    };

    var clipDuration      = customClipDuration !== null ? customClipDuration : 9; // seconds
    var blockedClipRanges = Array.isArray(options && options.blockedClipRanges) ? options.blockedClipRanges : [];

    if (blockedClipRanges.length) {
      logQYM('  [blocked clips] ' + blockedClipRanges.length + ' clip(s) excluded from placement: ' +
        blockedClipRanges.map(function (r) { return '"' + (r.name || '?') + '" (' + secToMS(r.start) + '–' + secToMS(r.end) + ')'; }).join(', '));
    }

    // Returns true when a candidate at time t (lasting clipDuration seconds)
    // would overlap any named clip that must not be covered.
    function overlapsBlockedClip(t) {
      var endT = t + clipDuration;
      for (var bi = 0; bi < blockedClipRanges.length; bi++) {
        var br = blockedClipRanges[bi];
        if (endT > br.start && t < br.end) return true;
      }
      return false;
    }

    // Candidate grid density — 5s step balances coverage vs. frame export count.
    var skipStep     = 5;
    var slideGap     = 0.1;
    var posCache     = {};
    var cursor       = windowStart;
    var index        = 0;
    var results      = [];
    var placedIntervals = [];   // {start,end} for every placed slide — used for gap-fill

    var step = slides.length > 1 ? (windowEnd - windowStart) / slides.length : 0;

    // Gap candidates between already-placed slides — same helper as the
    // Slides tab uses at line ~1746. Returns sorted candidate start-times
    // that fit `dur` seconds inside gaps between intervals.
    function gapCandidatesQYM(zoneSt, zoneEd, intervals, dur) {
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

    function previewCachedQYM(t, cb) {
      var key = Math.round(t);
      if (posCache[key] !== undefined) {
        var c = posCache[key];
        cb(null, c ? {
          startSeconds: t,
          resolvedAnchor: c.resolvedAnchor,
          allUnsafe:      c.allUnsafe,
          scores:         c.scores
        } : null);
        return;
      }
      // Reuse the existing previewSinglePlacement (exports frames from the active sequence)
      // We pass a synthetic placement object — only startSeconds matters for frame export
      previewSinglePlacement(options, { clipDurationSeconds: clipDuration, categoryName: 'QYM', language: '' }, t, function (err, result) {
        posCache[key] = (err || !result) ? null : {
          resolvedAnchor: result.resolvedAnchor,
          allUnsafe:      result.allUnsafe,
          scores:         result.scores
        };
        cb(err, result);
      });
    }

    function placeNext() {
      if (index >= slides.length) { callback(null, results); return; }

      var hardDeadline = windowEnd - clipDuration;
      var originalStart = windowStart + index * step;
      var target = Math.max(cursor, Math.min(originalStart, hardDeadline));
      if (target > hardDeadline) target = hardDeadline;

      var maxSteps = Math.max(Math.ceil((hardDeadline - target) / skipStep), Math.ceil((target - cursor) / skipStep), 1);
      maxSteps = Math.min(maxSteps, 30);  // cap: at most ~60 forward candidates per slide
      var candidates = [];
      var seen = {};
      var gapMarker = {};  // candidate-key → true if it's a gap insert (don't advance cursor)

      // Skip a candidate that overlaps any already-placed slide (Slides-tab rule
      // at line 2000). Each slide is `clipDuration` seconds; if t..t+clipDuration
      // intersects an existing slide's interval, scanning it is pointless.
      function overlapsPlaced(t) {
        var endT = t + clipDuration;
        for (var pi = 0; pi < placedIntervals.length; pi++) {
          var iv = placedIntervals[pi];
          if (endT > iv.start && t < iv.end) return true;
        }
        return false;
      }

      // Forward candidate (must respect cursor — sequential placement).
      function addC(t) {
        t = Math.round(t * 10) / 10;
        if (t < cursor - 0.05 || t > hardDeadline + 0.05) return;
        if (overlapsPlaced(t)) return;
        if (overlapsBlockedClip(t)) return;
        var k = Math.round(t); // 1s dedup buckets — 0.1s twins are identical for a 9s clip
        if (seen[k]) return;
        seen[k] = true;
        candidates.push(t);
      }

      // Gap candidate (allowed BEFORE cursor — fits in a hole left by an
      // earlier face-pushed slide). Same overlap and bounds checks.
      function addCGap(t) {
        t = Math.round(t * 10) / 10;
        if (t < windowStart - 0.05 || t > hardDeadline + 0.05) return;
        if (overlapsPlaced(t)) return;
        if (overlapsBlockedClip(t)) return;
        var k = Math.round(t); // 1s dedup buckets
        if (seen[k]) return;
        seen[k] = true;
        gapMarker[k] = true;
        candidates.push(t);
      }

      // ── Forward candidates (close-to-ideal time first) ───────────────
      addC(target);
      for (var s = 1; s <= maxSteps; s++) { addC(target + s * skipStep); addC(target - s * skipStep); }

      // ── Gap candidates from the start (Slides-tab pattern) ───────────
      // Earlier slides may have been face-pushed forward, leaving open
      // windows before the cursor. Mix gap candidates into the main list
      // so they're tried alongside forward ones, sorted by proximity.
      var gapsAtStart = gapCandidatesQYM(windowStart, windowEnd - clipDuration, placedIntervals, clipDuration);
      for (var gi = 0; gi < gapsAtStart.length; gi++) addCGap(gapsAtStart[gi]);

      // Single sort by proximity to the slide's ideal time
      candidates.sort(function (a, b) { return Math.abs(a - originalStart) - Math.abs(b - originalStart); });

      var gapCount = 0;
      for (var gci in gapMarker) if (gapMarker.hasOwnProperty(gci)) gapCount++;
      logQYM('Slide ' + (index + 1) + ' [' + slides[index].type + ']: searching ' + candidates.length + ' candidates (' + (candidates.length - gapCount) + ' forward + ' + gapCount + ' gap)…');

      var bestFallback  = null;
      var attemptIdx    = 0;
      var fineCands     = null;
      var fineIdx2      = 0;
      var rejectedFace  = 0;   // top-right blocked by a face
      var rejectedText  = 0;   // (legacy slot — text gate disabled for QYM, kept for log compat)
      var rejectedSkin  = 0;   // pixel-heuristic skin too high
      var rejectedBoth  = 0;   // (legacy slot — allUnsafe gate dropped for QYM)
      var traceLogged   = 0;   // capped detailed-trace counter
      var TRACE_CAP     = 60;  // log up to this many per-candidate evaluations per slide

      // Detailed trace log — coco-ssd terms.
      //   persons    = total persons detected anywhere in the frame
      //   in-TR/conf = a person box overlaps the slide region (with 3% margin)
      //                and its confidence
      //   dx         = 2D distance from the nearest person box to TR
      //                (0 = inside, 0.5 = far across the frame, ∞ = no people)
      function traceLog(t, isGap, cls, preview, fromCache) {
        if (traceLogged >= TRACE_CAP) return;
        traceLogged++;
        var src = isGap ? 'gap' : 'fwd';
        if (fromCache) src += '/cache';
        // Show DF timecode AND real seconds — the saved frames embed real
        // seconds in their filename (e.g. "_t555s79") so the user can match
        // log entries to debug frames unambiguously.
        var detail = '';
        if (preview && preview.scores) {
          var s = preview.scores['top-right'];
          if (s) {
            var bits = [];
            // Primary verdict: PERSON > TEXT > clear
            if (s.faceDetected) {
              bits.push('PERSON-IN-TR conf=' + (s.maxConfidence || 0).toFixed(2));
            } else if (s.textDetected) {
              bits.push('TEXT-IN-TR "' + (s.recognizedText || '').slice(0, 40) + '"');
            } else {
              bits.push('TR-clear');
            }
            if ((s.faceCount || 0) > 0) bits.push('persons=' + s.faceCount);
            // Show how many of the N sample frames flagged unsafe.
            if ((s.totalSampleCount || 0) > 0) {
              bits.push('samples=' + (s.unsafeSampleCount || 0) + '/' + s.totalSampleCount);
            }
            if (typeof s.nearestFaceDx === 'number') {
              if (isFinite(s.nearestFaceDx)) bits.push('dx=' + s.nearestFaceDx.toFixed(2));
              else                            bits.push('dx=∞');
            }
            detail = bits.join(' ');
          }
        } else {
          detail = '(no-data)';
        }
        var marker = (cls === 'safe') ? '✓ SAFE  ' : '✗ ' + cls.toUpperCase() + ' ';
        logQYM('    ' + marker + '[' + src + ' ' + secToMS(t) + ' (' + (Math.round(t * 100) / 100).toFixed(2) + 's)] ' + detail);
      }

      // QYM-specific: slides MUST go top-right. The frame analyzer also evaluates
      // top-left and may choose top-left as the safer corner — but we can't switch
      // corners. We only accept a candidate when the top-right corner itself is
      // clear (no face, no skin/head, no text overlap).
      //
      // We check skin/score even on the NN path because the face detector misses
      // profile views, faces from behind, and heads inside hijabs. The pixel
      // analysis (skin pixels, edges, texture) catches those.
      //
      // Returns 'safe' | 'face' | 'text' | 'skin' | 'both' | 'no-data'.
      //
      // Threshold tuning (verified against actual scoreAnchorRegion output):
      //   `score` is a 0–~75 weighted sum (skinRatio*25 + topSkin*20 + ...) —
      //   not a 0–1 fraction. The previous 0.30 gate fired on essentially every
      //   busy frame.  We now rely on skinRatio alone (a clean 0–1 fraction)
      //   plus textRatio plus the explicit face flag.
      //
      // ── Two-stage gate ─────────────────────────────────────────────────
      //
      //   1. Person check (coco-ssd):  any detected person overlaps TR → reject
      //   2. Text check (OCRAD):       any broadcast caption text in TR → reject
      //
      //   `tr.faceDetected` carries the person-overlap flag (legacy field name).
      //   `tr.textDetected` carries the OCR text-overlap flag (new).
      //   The OCR check only runs when the person check passes — see runCocoSsdAnalysis.
      function classifyQYM(preview) {
        if (!preview) return 'no-data';
        var scores = preview.scores;
        if (!scores) return 'safe';
        var tr = scores['top-right'];
        if (!tr) return 'no-data';
        if (tr.faceDetected) return 'face';   // coco-ssd: person box overlaps TR
        if (tr.textDetected) return 'text';   // OCRAD: broadcast caption in TR
        return 'safe';
      }
      function isTopRightSafeQYM(preview) {
        return classifyQYM(preview) === 'safe';
      }

      // Compact one-line description for the per-slide success / fallback log.
      function describeQYM(preview) {
        if (!preview || !preview.scores) return 'no-analysis';
        var tr = preview.scores['top-right'];
        if (!tr) return 'no-data';
        var bits = [];
        if (tr.faceDetected) {
          bits.push('PERSON-IN-TR(c=' + (tr.maxConfidence || 0).toFixed(2) + ')');
        } else if (tr.textDetected) {
          bits.push('TEXT-IN-TR("' + (tr.recognizedText || '').slice(0, 30) + '")');
        } else {
          bits.push('TR-clear');
        }
        if ((tr.faceCount || 0) > 0) {
          bits.push((tr.faceCount || 0) + ' person' + ((tr.faceCount === 1) ? '' : 's') + ' in frame');
        } else {
          bits.push('no people in frame');
        }
        return bits.join(' | ');
      }

      // Fallback ranking: prefer the candidate where TOP-RIGHT is least bad.
      function topRightScoreQYM(preview) {
        if (!preview || !preview.scores) return 9999;
        var tr = preview.scores['top-right'];
        if (!tr) return 9999;
        if (typeof tr.faceDetected !== 'undefined') {
          return (tr.faceDetected ? 100 : 0) + (tr.maxConfidence || 0) + (tr.textRatio || 0) * 5;
        }
        return (tr.score || 0) + (tr.skinRatio || 0) * 2;
      }

      function updateBestQYM(preview) {
        if (!preview) return;
        if (!bestFallback || topRightScoreQYM(preview) < topRightScoreQYM(bestFallback)) bestFallback = preview;
      }

      function bumpReject(reason) {
        if (reason === 'face') rejectedFace++;
        else if (reason === 'text') rejectedText++;
        else if (reason === 'skin') rejectedSkin++;
        else if (reason === 'both') rejectedBoth++;
      }

      function finalizeQYM(result, isFallback, gapInsert) {
        var rejSummary = 'face:' + rejectedFace + ' text:' + rejectedText +
                         ' skin:' + rejectedSkin + ' both:' + rejectedBoth;

        // ── Skip-on-fail ─────────────────────────────────────────────────
        // If we exhausted every candidate AND the best fallback still has a
        // detected face or significant skin in the top-right, the scene is
        // genuinely unsuited for this slide. Skip the slide entirely so the
        // viewer doesn't see an ad card landing on someone's face. The
        // language gets reverted in runQuanYinAndMax so the next run gets
        // another chance at it.
        if (isFallback) {
          var bad = false;
          if (!result || !result.scores) {
            bad = true; // no data at all — safer to skip than to drop a slide blind
          } else {
            var tr = result.scores['top-right'];
            if (!tr) bad = true;
            else if (tr.faceDetected) bad = true;  // coco-ssd: person overlaps TR
            else if (tr.textDetected) bad = true;  // OCRAD: caption overlaps TR
          }
          if (bad) {
            results.push({ skipped: true, reason: 'no-clear-position', describe: describeQYM(result) });
            logQYM('  ⏭ Slide ' + (index + 1) + ' SKIPPED — no clear position in window. Language will be reverted. [skipped ' + rejSummary + ']');
            // Cursor stays where it was — next slide can use this slide's would-be window.
            index++;
            placeNext();
            return;
          }
        }

        var finalStart  = (result && typeof result.startSeconds === 'number') ? result.startSeconds : target;
        // Only clamp to cursor for forward placement — gap inserts go BEFORE cursor on purpose.
        if (!gapInsert && finalStart < cursor) finalStart = cursor;
        if (finalStart > hardDeadline) finalStart = hardDeadline;
        finalStart = Math.max(finalStart, 0);

        results.push({ startSeconds: finalStart, resolvedAnchor: 'top-right' });
        placedIntervals.push({ start: finalStart, end: finalStart + clipDuration });

        if (isFallback) {
          logQYM('  ⚠ Slide ' + (index + 1) + ' fallback @ ' + secToMS(finalStart) +
                 ' (acceptable but not ideal) — ' + describeQYM(result) + ' [skipped ' + rejSummary + ']');
        } else if (gapInsert) {
          logQYM('  ↩ Slide ' + (index + 1) + ' gap-inserted @ ' + secToMS(finalStart) +
                 ' — ' + describeQYM(result) + ' [skipped ' + rejSummary + ']');
        } else {
          logQYM('  ✓ Slide ' + (index + 1) + ' @ ' + secToMS(finalStart) +
                 ' — ' + describeQYM(result) + ' [skipped ' + rejSummary + ']');
        }

        // Gap insert keeps cursor where it was — only forward placement advances it.
        if (!gapInsert) {
          cursor = finalStart + clipDuration + slideGap;
        }
        index++;
        placeNext();
      }

      // No early-bail — the Slides tab's buildSafePlacementPlan tries every
      // candidate before falling back, and so do we. The position cache makes
      // repeated lookups across slides nearly free, and giving up early was
      // missing valid gaps far from the slide's ideal time (e.g. open windows
      // mid-timeline that only show up after closer candidates fail).

      function tryNextQYM() {
        if (attemptIdx >= candidates.length) {
          // Forward + gap candidates exhausted — try a 1s fine pass as last resort
          if (!fineCands) {
            var fine = [];
            var ft = cursor;
            while (ft <= hardDeadline + 0.05) {
              var fk = Math.round(ft);
              if (!seen[fk] && !overlapsPlaced(ft) && !overlapsBlockedClip(ft)) fine.push(Math.round(ft * 10) / 10);
              ft += 1.0;
            }
            fine.sort(function (a, b) { return Math.abs(a - originalStart) - Math.abs(b - originalStart); });
            fineCands = fine.slice(0, 10);
          }
          if (fineIdx2 < fineCands.length) {
            var ft2 = fineCands[fineIdx2++];
            var fc2 = posCache[Math.round(ft2)] !== undefined;
            previewCachedQYM(ft2, function (err, preview) {
              if (!err && preview) {
                updateBestQYM(preview);
                var cls2 = classifyQYM(preview);
                traceLog(ft2, false, cls2, preview, fc2);
                if (cls2 === 'safe') { finalizeQYM(preview, false); return; }
                bumpReject(cls2);
              } else {
                traceLog(ft2, false, 'no-data', null, fc2);
              }
              tryNextQYM();
            });
            return;
          }
          // Everything tried — use bestFallback (skip-on-fail logic in finalizeQYM
          // will decide whether to actually drop the slide).
          finalizeQYM(bestFallback, true);
          return;
        }
        var t = candidates[attemptIdx++];
        var isGap = !!gapMarker[Math.round(t)];
        var fromCache = posCache[Math.round(t)] !== undefined;
        previewCachedQYM(t, function (err, preview) {
          if (!err && preview) {
            updateBestQYM(preview);
            var cls = classifyQYM(preview);
            traceLog(t, isGap, cls, preview, fromCache);
            if (cls === 'safe') { finalizeQYM(preview, false, isGap); return; }
            bumpReject(cls);
          } else {
            traceLog(t, isGap, 'no-data', null, fromCache);
          }
          tryNextQYM();
        });
      }

      tryNextQYM();
    }

    placeNext();
  }

  // ── Main QYM run ─────────────────────────────────────────────────────────
  function runQuanYinAndMax() {
    if (qymRunning) return;

    // Read everything up-front so we can validate before disabling the UI.
    var slideType    = getQymSlideType();   // 'both' | 'qy-only' | 'max-only'
    var needQY       = (slideType === 'both' || slideType === 'qy-only');
    var needMax      = (slideType === 'both' || slideType === 'max-only');
    var qyFolder     = selectedQYFolder;
    var maxFolder    = selectedMaxFolder;
    var maxCount     = parseInt(qymMaxCountInput.value, 10)    || 4;
    var targetTrack  = parseInt(qymTargetTrackInput.value, 10) || 10;
    var fadeDuration = parseFloat(qymFadeDurationInput.value)  || 0.3;
    var avoidFaces   = !!(qymAvoidFacesInput && qymAvoidFacesInput.checked);
    var clipDuration = 9;

    if (needQY && (!qyFolder || !fs.existsSync(qyFolder))) {
      setQymStatus('Error: Quan-Yin folder not set or not found.');
      return;
    }
    if (needMax && (!maxFolder || !fs.existsSync(maxFolder))) {
      setQymStatus('Error: SMTV Max folder not set or not found.');
      return;
    }

    // Lock UI immediately — before any work, including the JSX window probe.
    setQymRunning(true);
    setQymStatus('Working…');

    // Initialise debug-frame folder if the option is enabled. Done before
    // anything analyzes a frame, so every save lands in the same run folder.
    var debugDir = _qymInitDebugDir();
    if (debugDir) {
      logQYM('  [debug-frames] Saving annotated frames to: ' + debugDir);
    }

    // 1. Probe sequence window first so we can warn if no In/Out is set.
    callJsx('qymGetSequenceWindow', {}, function (rawWin) {
      var win;
      try { win = JSON.parse(rawWin); } catch (e) {
        setQymStatus('Error: Failed to read sequence window — ' + rawWin);
        setQymRunning(false); return;
      }
      if (!win.ok) {
        setQymStatus('Error: ' + (win.error || 'Could not read sequence.'));
        setQymRunning(false); return;
      }

      var windowStart       = win.windowStart;
      var windowEnd         = win.windowEnd;
      var usedLen           = win.usedTimelineLengthSeconds;
      var blockedClipRanges = Array.isArray(win.blockedClipRanges) ? win.blockedClipRanges : [];
      // Heuristic: if the placement window is essentially the whole used
      // timeline (start ≈ 0 and end ≈ used length, ±0.5s slop), there are
      // no In/Out points active. Warn the user.
      var hasInOut    = (windowStart > 0.5) || (windowEnd < usedLen - 0.5);

      // Log and subtract blocked clip ranges from usable window length.
      var blockedInWindow = 0;
      for (var bci = 0; bci < blockedClipRanges.length; bci++) {
        var br = blockedClipRanges[bci];
        var overlapStart = Math.max(br.start, windowStart);
        var overlapEnd   = Math.min(br.end,   windowEnd);
        if (overlapEnd > overlapStart) blockedInWindow += overlapEnd - overlapStart;
        logQYM('⛔ Blocked clip: "' + (br.name || '?') + '" ' + secToMS(br.start) + '–' + secToMS(br.end));
      }

      // ── CAPACITY CAP ─────────────────────────────────────────────────────
      // Calculate the max number of 9-second slides that physically fit in
      // the window BEFORE we pick languages or run frame analysis. Saves
      // burning languages from the cycle on slides that won't be placed,
      // and avoids minutes of wasted person-detection on slides that would
      // be capped anyway.
      var windowLen   = windowEnd - windowStart;
      var usableLen   = Math.max(0, windowLen - blockedInWindow);
      var capacity    = Math.floor(usableLen / (clipDuration + 0.1));
      var typesNeeded = (needQY ? 1 : 0) + (needMax ? 1 : 0);
      var requested   = maxCount * typesNeeded;

      if (capacity < typesNeeded) {
        // Window too small for a full-length slide of each requested type.
        //
        // Special case (single-type mode + window ≥ 2s): drop to 1 slide
        // and SHORTEN it to fit the window. e.g. an 8-second window in
        // qy-only mode → 1 Quan-Yin slide of 8 seconds instead of giving up.
        var MIN_SHORT_SLIDE = 2.0;
        if (typesNeeded === 1 && windowLen >= MIN_SHORT_SLIDE) {
          var shortDuration = Math.max(MIN_SHORT_SLIDE, windowLen - 0.1);
          logQYM('⚠ Window is ' + windowLen.toFixed(1) + 's — too short for a standard ' + clipDuration + 's slide. Fitting 1 shortened slide of ' + shortDuration.toFixed(1) + 's.');
          maxCount     = 1;
          clipDuration = shortDuration;   // override for this run only
        } else {
          setQymStatus('Error: window is ' + windowLen.toFixed(1) + 's — too short to fit even one slide' + (typesNeeded > 1 ? ' per selected type' : '') + ' (need ' + (typesNeeded * (clipDuration + 0.1)).toFixed(1) + 's' + (typesNeeded > 1 ? '; switch to Quan-Yin only or SMTV Max only for shorter windows' : '') + ').');
          setQymRunning(false);
          return;
        }
      } else if (requested > capacity) {
        var newMaxCount = Math.max(1, Math.floor(capacity / typesNeeded));
        var capDescr;
        if (typesNeeded === 1) {
          capDescr = capacity + ' slide' + (capacity === 1 ? '' : 's');
        } else {
          capDescr = newMaxCount + ' per type (' + (newMaxCount * typesNeeded) + ' total)';
        }
        logQYM('⚠ Window is ' + windowLen.toFixed(1) + 's — only ' + capDescr + ' fit. You requested ' +
               requested + (typesNeeded > 1 ? ' (' + maxCount + ' per type)' : '') +
               '. Capping to ' + newMaxCount + (typesNeeded > 1 ? ' per type.' : '.'));
        maxCount = newMaxCount;
      }

      var proceed = function () { runQYMPlacement(slideType, needQY, needMax, qyFolder, maxFolder, maxCount, targetTrack, fadeDuration, avoidFaces, clipDuration, windowStart, windowEnd, blockedClipRanges); };

      if (!hasInOut) {
        showQymNoInOutConfirm(function (confirmed) {
          if (!confirmed) {
            setQymStatus('Cancelled — set In/Out points and try again.');
            setQymRunning(false);
            return;
          }
          logQYM('⚠ No In/Out points set — placing across the whole sequence.');
          proceed();
        });
      } else {
        proceed();
      }
    });
  }

  // Inner placement runner — entered after the sequence window has been
  // probed and the user has confirmed (or didn't need to confirm) running
  // without In/Out points.
  function runQYMPlacement(slideType, needQY, needMax, qyFolder, maxFolder, maxCount, targetTrack, fadeDuration, avoidFaces, clipDuration, windowStart, windowEnd, blockedClipRanges) {
    try {
      // 2. Scan folders we actually need
      var qyScan  = needQY  ? scanQuanYinFolder(qyFolder)  : { englishA: null, englishB: null, nonEnglish: [] };
      var maxScan = needMax ? scanSmtvMaxFolder(maxFolder) : { english:  null,                  nonEnglish: [] };

      if (needQY && !qyScan.englishA && !qyScan.englishB) {
        setQymStatus('Error: No English file found in Quan-Yin folder.');
        setQymRunning(false); return;
      }
      if (needMax && !maxScan.english) {
        setQymStatus('Error: No English file (SMTV Max-ENG.png) found in SMTV Max folder.');
        setQymRunning(false); return;
      }

      // 3. Pick slides with cycle tracking. Pick Quan-Yin first (when in use)
      // so SMTV Max can avoid overlapping languages.
      var tracking = loadTracking();
      var qyResult  = needQY  ? pickQuanYinSlides(qyScan, maxCount, tracking) : { files: [], isEnglish: [], langs: [], warnings: [] };

      var qyCanonExclude = {};
      var qyCanonPairs   = [];
      qyResult.langs.forEach(function (lang, idx) {
        if (idx === 0) return; // skip the English entry
        var canon = _qymCanonLang(lang);
        if (canon) {
          qyCanonExclude[canon] = true;
          qyCanonPairs.push(lang + '→' + canon);
        }
      });

      var maxResult = needMax ? pickSmtvMaxSlides(maxScan, maxCount, tracking, qyCanonExclude) : { files: [], isEnglish: [], langs: [], warnings: [] };
      saveTracking(tracking); // persist cycle state immediately

      // Log warnings
      qyResult.warnings.concat(maxResult.warnings).forEach(function (w) { logQYM('⚠ ' + w); });

      // Log selection summary (only the types we're actually using)
      if (needQY) {
        logQYM('Quan-Yin: using English ' + (tracking.quanYinCycle.nextEnglish === 'A' ? 'B' : 'A') +
               ' + ' + Math.max(0, qyResult.files.length - 1) + ' non-English: ' +
               qyResult.langs.slice(1).join(', '));
      }
      if (needMax) {
        logQYM('SMTV Max: 1 English + ' + Math.max(0, maxResult.files.length - 1) + ' non-English: ' +
               maxResult.langs.slice(1).join(', '));
      }
      if (slideType !== 'both') {
        logQYM('  [mode] ' + (slideType === 'qy-only' ? 'Quan-Yin only — SMTV Max skipped' : 'SMTV Max only — Quan-Yin skipped'));
      }

      // Diagnostic: canonical mapping (only when both types are in play —
      // overlap avoidance is meaningless otherwise).
      if (needQY && needMax) {
        var maxCanonPairs = [];
        maxResult.langs.forEach(function (lang, idx) {
          if (idx === 0) return;
          maxCanonPairs.push(lang + '→' + _qymCanonLang(lang));
        });
        logQYM('  [canon] QY:  ' + qyCanonPairs.join(', '));
        logQYM('  [canon] Max: ' + maxCanonPairs.join(', '));
        var overlaps = [];
        maxResult.langs.forEach(function (lang, idx) {
          if (idx === 0) return;
          var c = _qymCanonLang(lang);
          if (qyCanonExclude[c]) overlaps.push(lang + '(' + c + ')');
        });
        if (overlaps.length) logQYM('  ⚠ overlap with Quan-Yin: ' + overlaps.join(', '));
      }

      // 4. Build slide list — interleave when both types present, otherwise
      // emit only the picked type.
      var slides;
      if (needQY && needMax) {
        slides = interleaveQYM(qyResult, maxResult);
      } else if (needQY) {
        slides = qyResult.files.map(function (fp, i) {
          return { filePath: fp, type: 'quan-yin', isEnglish: qyResult.isEnglish[i], lang: qyResult.langs[i] };
        });
      } else {
        slides = maxResult.files.map(function (fp, i) {
          return { filePath: fp, type: 'smtv-max', isEnglish: maxResult.isEnglish[i], lang: maxResult.langs[i] };
        });
      }
      if (!slides.length) {
        setQymStatus('Error: No slides selected.');
        setQymRunning(false); return;
      }

      // 5. Capacity check & placement (uses the windowStart/windowEnd we
      // already probed at the top of runQuanYinAndMax).
      runQYMPlacementWithWindow(slides, tracking, windowStart, windowEnd, targetTrack, fadeDuration, avoidFaces, clipDuration, blockedClipRanges);
    } catch (err) {
      setQymStatus('Error: ' + (err.message || err));
      setQymRunning(false);
    }
  }

  // Reverts language tracking for a list of slides that didn't actually
  // make it onto the timeline. Mutates `tracking` in place; returns a small
  // summary so the caller can log it. Saving is the caller's responsibility.
  //
  //   - non-English slides: removed from their cycle's usedNonEnglish[]
  //   - English slides:     nextEnglish toggle inverted (so same variant
  //                         comes up again next run)
  function _qymRevertCycleForSlides(slidesToRevert, tracking) {
    var revertedQY = [], revertedMax = [], toggled = false;
    for (var i = 0; i < slidesToRevert.length; i++) {
      var s = slidesToRevert[i];
      if (!s) continue;
      if (s.type === 'quan-yin') {
        if (!s.isEnglish && tracking.quanYinCycle && Array.isArray(tracking.quanYinCycle.usedNonEnglish)) {
          var qIdx = tracking.quanYinCycle.usedNonEnglish.indexOf(s.lang);
          if (qIdx !== -1) {
            tracking.quanYinCycle.usedNonEnglish.splice(qIdx, 1);
            revertedQY.push(s.lang);
          }
        }
        if (s.isEnglish && tracking.quanYinCycle) {
          tracking.quanYinCycle.nextEnglish = (tracking.quanYinCycle.nextEnglish === 'A') ? 'B' : 'A';
          toggled = true;
        }
      } else if (s.type === 'smtv-max') {
        if (!s.isEnglish && tracking.smtvMaxCycle && Array.isArray(tracking.smtvMaxCycle.usedNonEnglish)) {
          var mIdx = tracking.smtvMaxCycle.usedNonEnglish.indexOf(s.lang);
          if (mIdx !== -1) {
            tracking.smtvMaxCycle.usedNonEnglish.splice(mIdx, 1);
            revertedMax.push(s.lang);
          }
        }
      }
    }
    return { revertedQY: revertedQY, revertedMax: revertedMax, toggled: toggled };
  }

  // Logs the revert summary and saves tracking if anything actually changed.
  function _qymCommitRevert(summary, tracking) {
    if (summary.revertedQY.length)  logQYM('  ↺ Reverted Quan-Yin languages: ' + summary.revertedQY.join(', '));
    if (summary.revertedMax.length) logQYM('  ↺ Reverted SMTV Max languages: ' + summary.revertedMax.join(', '));
    if (summary.toggled)            logQYM('  ↺ Reverted Quan-Yin English A/B toggle.');
    if (summary.revertedQY.length || summary.revertedMax.length || summary.toggled) {
      saveTracking(tracking);
    }
  }

  function runQYMPlacementWithWindow(slides, tracking, windowStart, windowEnd, targetTrack, fadeDuration, avoidFaces, clipDuration, blockedClipRanges) {
    blockedClipRanges = Array.isArray(blockedClipRanges) ? blockedClipRanges : [];
    var windowLen = windowEnd - windowStart;
    // Capacity check
    var maxFit = Math.floor(windowLen / (clipDuration + 0.1));
    if (slides.length > maxFit) {
      logQYM('⚠ Window too short for ' + slides.length + ' slides — capping to ' + maxFit + '.');
      slides = slides.slice(0, maxFit);
    }
    if (!slides.length) {
      setQymStatus('Error: sequence window is too short to fit even one slide.');
      setQymRunning(false); return;
    }

    logQYM('Window: ' + secToMS(windowStart) + ' – ' + secToMS(windowEnd) +
           ' | ' + slides.length + ' slides × ' + clipDuration + 's | avoidFaces=' + avoidFaces);

    function doPlace(placements) {
      var payload = {
        slides:      slides,
        placements:  placements,
        targetTrack: targetTrack,
        clipDuration: clipDuration,
        fadeDuration: fadeDuration
      };
      callJsx('quanYinMaxImportAndPlace', payload, function (rawRes) {
        var res;
        try { res = JSON.parse(rawRes); } catch (e) {
          // JSX returned non-JSON — treat as total failure, revert ALL slides
          logQYM('  Placement failed before any clip was imported — reverting all language picks.');
          _qymCommitRevert(_qymRevertCycleForSlides(slides, tracking), tracking);
          setQymStatus('Error: ' + rawRes);
          setQymRunning(false); return;
        }
        if (!res.ok) {
          // JSX reported failure — nothing got placed, revert ALL slides
          logQYM('  Placement reported failure — reverting all language picks.');
          _qymCommitRevert(_qymRevertCycleForSlides(slides, tracking), tracking);
          setQymStatus('Error: ' + (res.error || 'Placement failed.'));
          setQymRunning(false); return;
        }
        // JSX OK but reported placedCount < slides.length — some clips
        // didn't import (project-item lookup miss, etc.). If ZERO placed,
        // revert everything; otherwise log the gap (we can't know which
        // specific slides failed from JSX's response).
        if ((res.placedCount || 0) === 0 && slides.length > 0) {
          logQYM('  ⚠ JSX reported 0 slides placed — reverting all language picks.');
          _qymCommitRevert(_qymRevertCycleForSlides(slides, tracking), tracking);
          setQymStatus('Error: No slides were placed (all imports failed).');
          setQymRunning(false); return;
        }
        if (res.placedCount < slides.length) {
          logQYM('  ⚠ ' + (slides.length - res.placedCount) + ' slide(s) did not import — languages may be over-counted in cycle (rerun to verify).');
        }
        logQYM('✓ Placed ' + res.placedCount + ' slides on track ' + targetTrack + '.');
        setQymStatus(qymStatusEl.textContent.trim());
        setQymRunning(false);
      });
    }

    if (avoidFaces) {
      logQYM('Checking frames for faces/heads…');
      buildSafePlacementPlanQYM(
        { slideAnchor: 'top-right', avoidFaces: true, clipDuration: clipDuration, blockedClipRanges: blockedClipRanges },
        slides, windowStart, windowEnd,
        function (err, placements) {
          if (err || !placements) {
            logQYM('Face check failed (' + (err ? err.message : 'unknown') + ') — using equal spacing.');
            doPlace(computeQYMPositions(slides, windowStart, windowEnd, clipDuration));
            return;
          }

          // Filter out skipped slides; revert their cycle state via helper.
          var keptSlides     = [];
          var keptPlacements = [];
          var skippedSlides  = [];
          for (var i = 0; i < placements.length; i++) {
            var p = placements[i];
            var s = slides[i];
            if (p && p.skipped) {
              skippedSlides.push(s);
              continue;
            }
            keptSlides.push(s);
            keptPlacements.push(p);
          }
          if (skippedSlides.length) {
            _qymCommitRevert(_qymRevertCycleForSlides(skippedSlides, tracking), tracking);
            logQYM('Skipped ' + skippedSlides.length + ' slide(s) due to no clear position. Placing ' + keptPlacements.length + '.');
          }
          if (!keptPlacements.length) {
            // Belt-and-suspenders: in case the loop above missed anything,
            // explicitly revert any remaining slides too.
            _qymCommitRevert(_qymRevertCycleForSlides(slides, tracking), tracking);
            setQymStatus('All slides skipped — no clear positions found in this sequence. Languages returned to cycle for next run.');
            setQymRunning(false);
            return;
          }
          slides = keptSlides;
          doPlace(keptPlacements);
        }
      );
    } else {
      doPlace(computeQYMPositions(slides, windowStart, windowEnd, clipDuration));
    }
  }


  // ── QYM folder browse helpers ────────────────────────────────────────────
  function setSelectedQYFolder(folderPath) {
    selectedQYFolder = folderPath;
    if (qymQYFolderInput) qymQYFolderInput.value = folderPath;
    persistQYMSettings();
  }
  function setSelectedMaxFolder(folderPath) {
    selectedMaxFolder = folderPath;
    if (qymMaxFolderInput) qymMaxFolderInput.value = folderPath;
    persistQYMSettings();
  }

  // ── QYM browse buttons ───────────────────────────────────────────────────
  if (qymQYBrowseBtn) {
    qymQYBrowseBtn.addEventListener('click', function () {
      try {
        if (window.cep && window.cep.fs && typeof window.cep.fs.showOpenDialogEx === 'function') {
          var r = window.cep.fs.showOpenDialogEx(false, true, 'Select the Quan-Yin slides folder');
          if (r && r.data && r.data.length) { setSelectedQYFolder(r.data[0]); return; }
        }
      } catch (e) {}
      if (qymQYPickerInput) qymQYPickerInput.click();
    });
  }
  if (qymQYPickerInput) {
    qymQYPickerInput.addEventListener('change', function (evt) {
      var files = evt.target.files;
      if (!files || !files.length) return;
      var rel = files[0].webkitRelativePath || '';
      var top = rel.split('/')[0];
      var abs = files[0].path;
      var dir = path.dirname(abs);
      while (path.basename(dir) !== top && dir !== path.dirname(dir)) dir = path.dirname(dir);
      setSelectedQYFolder(dir);
    });
  }
  if (qymMaxBrowseBtn) {
    qymMaxBrowseBtn.addEventListener('click', function () {
      try {
        if (window.cep && window.cep.fs && typeof window.cep.fs.showOpenDialogEx === 'function') {
          var r = window.cep.fs.showOpenDialogEx(false, true, 'Select the SMTV Max slides folder');
          if (r && r.data && r.data.length) { setSelectedMaxFolder(r.data[0]); return; }
        }
      } catch (e) {}
      if (qymMaxPickerInput) qymMaxPickerInput.click();
    });
  }
  if (qymMaxPickerInput) {
    qymMaxPickerInput.addEventListener('change', function (evt) {
      var files = evt.target.files;
      if (!files || !files.length) return;
      var rel = files[0].webkitRelativePath || '';
      var top = rel.split('/')[0];
      var abs = files[0].path;
      var dir = path.dirname(abs);
      while (path.basename(dir) !== top && dir !== path.dirname(dir)) dir = path.dirname(dir);
      setSelectedMaxFolder(dir);
    });
  }

  // ── QYM settings change listeners ───────────────────────────────────────
  if (qymMaxCountInput)     qymMaxCountInput.addEventListener('change',     persistQYMSettings);
  if (qymTargetTrackInput)  qymTargetTrackInput.addEventListener('change',  persistQYMSettings);
  if (qymFadeDurationInput) qymFadeDurationInput.addEventListener('change', persistQYMSettings);
  if (qymAvoidFacesInput)   qymAvoidFacesInput.addEventListener('change',   persistQYMSettings);
  if (qymKeepDebugInput)    qymKeepDebugInput.addEventListener('change',    persistQYMSettings);
  if (qymSlideTypeRadios) {
    for (var _qstIdx = 0; _qstIdx < qymSlideTypeRadios.length; _qstIdx++) {
      qymSlideTypeRadios[_qstIdx].addEventListener('change', persistQYMSettings);
    }
  }

  // ── QYM run button ───────────────────────────────────────────────────────
  if (qymRunBtn) {
    qymRunBtn.addEventListener('click', function () { runQuanYinAndMax(); });
  }

  // ════════════════════════════════════════════════════════════════════════

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
  if (packPerCategoryInput) packPerCategoryInput.addEventListener('change', persistSettings);
  installUpdateBtn.addEventListener('click', installLatestUpdate);
  window.addEventListener('beforeunload', persistSettings);

  runBtn.addEventListener('click', function () {
    setRunning(true);
    setStatus('Working...');

    if (!selectedRootFolder) {
      setStatus('Please choose the root folder that contains your slide files first.');
      setRunning(false);
      return;
    }

    var requestedCount = parseInt(slideCountInput.value, 10);
    var targetTrack = parseInt(targetTrackInput.value, 10);
    var ignoreV1 = !!ignoreV1Input.checked;
    var slideAnchor = String(slideAnchorInput.value || 'top-right');
    var avoidFaces = !!avoidFacesInput.checked;
    var packPerCategory = !!(packPerCategoryInput && packPerCategoryInput.checked);

    if (!requestedCount || requestedCount < 1) {
      setStatus('Please enter a valid number of slides.');
      setRunning(false);
      return;
    }
    if (!targetTrack || targetTrack < 1) {
      setStatus('Please enter a valid target video track number.');
      setRunning(false);
      return;
    }

    // Yield to the browser so it can repaint the grayed-out button before
    // the synchronous heavy work (scanAllCategories etc.) blocks the JS thread.
    setTimeout(function () {
      try {
        persistSettings();
        var tracking = loadTracking();
        tracking._currentRunUsedLanguages = [];
        var categoryDataList = scanAllCategories(selectedRootFolder);
        if (!categoryDataList.length) {
          setStatus('No usable slide files were found under the selected root folder.');
          setRunning(false);
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
          avoidFaces: avoidFaces,
          packPerCategory: packPerCategory
        });
      } catch (err) {
        setStatus('Error: ' + err.message);
        setRunning(false);
      }
    }, 0);
  });

  extensionRoot = resolveExtensionRoot();
  manifestPath = extensionRoot ? path.join(extensionRoot, 'CSXS', 'manifest.xml') : path.join(path.resolve(__dirname, '..'), 'CSXS', 'manifest.xml');
  restoreSettings();
  restoreQYMSettings();
  // Show the loading banner immediately so the user knows the panel is alive
  // while the (large) detection models initialise. Each init function calls
  // _markModelReady / _markModelFailed; the banner hides when all settle.
  _updateLoadingBanner();
  _updateRunButtonsForLoading();
  initFaceApi();    // Load TinyFaceDetector model (async — Slides tab uses this)
  initCocoSsd();    // Load COCO-SSD model (async — QYM tab uses this for person detection)
  initOCRAD();      // Load OCRAD.js OCR — synchronous, no Workers, no SharedArrayBuffer
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
