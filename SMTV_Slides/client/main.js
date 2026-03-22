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
  var updateRepoInput = document.getElementById('updateRepo');
  var installedVersionEl = document.getElementById('installedVersion');
  var latestVersionEl = document.getElementById('latestVersion');
  var checkUpdatesBtn = document.getElementById('checkUpdatesBtn');
  var installUpdateBtn = document.getElementById('installUpdateBtn');
  var updateStatusEl = document.getElementById('updateStatus');
  var statusEl = document.getElementById('status');
  var chosenTitleEl = document.getElementById('chosenTitle');
  var chosenLanguagesEl = document.getElementById('chosenLanguages');

  var selectedRootFolder = '';
  var extensionRoot = path.resolve(__dirname, '..');
  var manifestPath = path.join(extensionRoot, 'CSXS', 'manifest.xml');
  var trackingDir = path.join(os.homedir(), '.new-peace-maker');
  var trackingFile = path.join(trackingDir, 'usage-history.json');
  var categoryOrder = ['NEW PEACE MAKER', 'Be Vegan Keep Peace', 'Forgiveness', 'Save the Earth', 'Veganism'];
  var ignoredFolderNames = { 'AFTERCODECS HAP ALPHA': true };
  var updateState = {
    installedVersion: '',
    latestVersion: '',
    latestRelease: null,
    checking: false,
    installing: false
  };
  var autoUpdateCheckIntervalMs = 12 * 60 * 60 * 1000;

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
        updateRepo: '',
        lastUpdateCheckAt: '',
        lastAvailableVersion: '',
      }
    };
  }

  function log(msg) {
    statusEl.textContent += '\n' + msg;
    statusEl.scrollTop = statusEl.scrollHeight;
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
      if (typeof parsed.settings.updateRepo === 'undefined') parsed.settings.updateRepo = base.settings.updateRepo;
      if (typeof parsed.settings.lastUpdateCheckAt === 'undefined') parsed.settings.lastUpdateCheckAt = base.settings.lastUpdateCheckAt;
      if (typeof parsed.settings.lastAvailableVersion === 'undefined') parsed.settings.lastAvailableVersion = base.settings.lastAvailableVersion;
      return parsed;
    } catch (e) {
      return defaultTracking();
    }
  }

  function saveTracking(data) {
    ensureTrackingFile();
    fs.writeFileSync(trackingFile, JSON.stringify(data, null, 2), 'utf8');
  }

  function persistSettings() {
    var tracking = loadTracking();
    tracking.settings.rootFolder = selectedRootFolder || '';
    tracking.settings.slideCount = parseInt(slideCountInput.value, 10) || 6;
    tracking.settings.targetTrack = parseInt(targetTrackInput.value, 10) || 9;
    tracking.settings.ignoreV1 = !!(ignoreV1Input && ignoreV1Input.checked);
    tracking.settings.updateRepo = updateRepoInput ? String(updateRepoInput.value || '').trim() : '';
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
    updateRepoInput.value = tracking.settings.updateRepo || '';
  }

  function setUpdateStatus(msg) {
    updateStatusEl.textContent = msg;
  }

  function setUpdateUiState() {
    installedVersionEl.textContent = updateState.installedVersion || '-';
    latestVersionEl.textContent = updateState.latestVersion || '-';
    checkUpdatesBtn.disabled = updateState.checking || updateState.installing;
    installUpdateBtn.disabled = updateState.checking || updateState.installing || !updateState.latestRelease;
    installUpdateBtn.hidden = !updateState.latestRelease || compareVersions(updateState.latestVersion, updateState.installedVersion) <= 0;
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
    tracking.settings.updateRepo = String(updateRepoInput.value || '').trim();
    saveTracking(tracking);
  }

  function shouldBackgroundCheck() {
    var tracking = loadTracking();
    var repo = String(updateRepoInput.value || '').trim();
    if (!repo) return false;
    if (!tracking.settings.lastUpdateCheckAt) return true;
    var lastCheckTime = Date.parse(tracking.settings.lastUpdateCheckAt);
    if (!lastCheckTime) return true;
    return (Date.now() - lastCheckTime) >= autoUpdateCheckIntervalMs;
  }

  function getGitHubReleaseApiUrl(repo) {
    return 'https://api.github.com/repos/' + repo + '/releases/latest';
  }

  function getTempPath(name) {
    return path.join(os.tmpdir(), 'smtv-slides-updater', String(name || ''));
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
    var repo = String(updateRepoInput.value || '').trim();
    if (!repo) {
      updateState.latestRelease = null;
      updateState.latestVersion = '';
      setUpdateStatus('Enter a GitHub repo like owner/repo to enable update checks.');
      setUpdateUiState();
      return;
    }

    updateState.checking = true;
    setUpdateStatus(opts.silent ? 'Checking for updates in background...' : 'Checking for updates...');
    setUpdateUiState();

    requestJson(getGitHubReleaseApiUrl(repo), function (err, release) {
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
      updateState.latestVersion = latestVersion || '';
      updateState.latestRelease = zipAsset ? release : null;
      persistUpdateInfo(latestVersion);

      if (!zipAsset) {
        setUpdateStatus('A release was found, but no zip asset is available to install.');
      } else if (compareVersions(latestVersion, updateState.installedVersion) > 0) {
        setUpdateStatus('Version ' + latestVersion + ' is available. Click Update Now to install it.');
      } else {
        setUpdateStatus('You are up to date.');
      }
      setUpdateUiState();
    });
  }

  function installLatestUpdate() {
    var repo = String(updateRepoInput.value || '').trim();
    if (!repo) {
      setUpdateStatus('Enter a GitHub repo before installing updates.');
      return;
    }

    var release = updateState.latestRelease;
    var latestVersion = updateState.latestVersion;
    var zipAsset = getReleaseZipAsset(release);
    if (!release || !zipAsset) {
      setUpdateStatus('No downloadable update is ready yet. Run Check for Updates first.');
      return;
    }

    if (!window.confirm('Install version ' + latestVersion + ' from ' + repo + ' now? Premiere Pro should be restarted after the update.')) {
      return;
    }

    updateState.installing = true;
    setUpdateStatus('Downloading update...');
    setUpdateUiState();

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

        setUpdateStatus('Installing update...');
        installExtractedExtension(extractedExtensionRoot);
        updateState.installedVersion = readManifestVersion(manifestPath) || latestVersion;
        updateState.installing = false;
        updateState.latestRelease = compareVersions(updateState.latestVersion, updateState.installedVersion) > 0 ? updateState.latestRelease : null;
        setUpdateStatus('Update installed. Please restart Premiere Pro to load version ' + updateState.installedVersion + '.');
        setUpdateUiState();
      } catch (installErr) {
        updateState.installing = false;
        setUpdateStatus('Update install failed: ' + installErr.message);
        setUpdateUiState();
      }
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

  var canonicalLanguageAliasMap = {
    English: ['english', 'eng'],
    Arabic: ['arabic', 'ara'],
    Aulac: ['aulac', 'au', 'aul'],
    Bulgarian: ['bulgarian', 'bul'],
    Chinese: ['chinese', 'chi', 'zho', 'zh'],
    Croatian: ['croatian', 'cro', 'hrv'],
    Czech: ['czech', 'cze', 'ces'],
    Estonian: ['estonian', 'est'],
    Ewe: ['ewe'],
    French: ['french', 'fre', 'fra'],
    German: ['german', 'ger', 'deu'],
    Hebrew: ['hebrew', 'heb'],
    Hindi: ['hindi', 'hin'],
    Hungarian: ['hungarian', 'hun'],
    Indonesian: ['indonesian', 'ind', 'ina'],
    Italian: ['italian', 'ita'],
    Japanese: ['japanese', 'jap', 'jpn'],
    Korean: ['korean', 'kor'],
    Malayalam: ['malayalam', 'mal'],
    Mongolian: ['mongolian', 'mon'],
    Persian: ['persian', 'per', 'fas'],
    Polish: ['polish', 'pol'],
    Portuguese: ['portuguese', 'por'],
    Punjabi: ['punjabi', 'pun', 'pan'],
    Romanian: ['romanian', 'rom', 'ron'],
    Russian: ['russian', 'rus'],
    Serbian: ['serbian', 'srp', 'scc'],
    Spanish: ['spanish', 'spa'],
    Swedish: ['swedish', 'swe'],
    Thai: ['thai', 'tha'],
    Urdu: ['urdu', 'urd'],
    Zulu: ['zulu', 'zul']
  };

  function titleCaseWords(str) {
    return String(str || '').replace(/\b[a-z]/g, function (ch) { return ch.toUpperCase(); });
  }

  function canonicalizeLanguageName(languageName) {
    var langNorm = normalizeToken(languageName);
    var compactNorm = langNorm.replace(/\s+/g, '');
    if (!langNorm) return null;

    for (var canonicalName in canonicalLanguageAliasMap) {
      if (!canonicalLanguageAliasMap.hasOwnProperty(canonicalName)) continue;
      var aliases = canonicalLanguageAliasMap[canonicalName];
      for (var i = 0; i < aliases.length; i++) {
        var alias = aliases[i];
        if (langNorm === alias || compactNorm === alias.replace(/\s+/g, '')) {
          return canonicalName;
        }
      }
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

  function scanSingleCategory(categoryName, categoryRoot) {
    if (categoryName === 'Be Vegan Keep Peace') {
      var flatFiles = fs.readdirSync(categoryRoot, { withFileTypes: true })
        .filter(function (d) { return d.isFile() && path.extname(d.name).toLowerCase() === '.mov'; })
        .map(function (d) { return path.join(categoryRoot, d.name); });

      var flatTitle = 'Be Vegan Keep Peace';
      var flatTitlesMap = {};
      flatTitlesMap[flatTitle] = {};

      flatFiles.forEach(function (fullPath) {
        var lang = parseFlatBeVeganLanguage(fullPath);
        if (!lang) return;
        flatTitlesMap[flatTitle][lang] = fullPath;
      });

      return {
        isFlatSingleTitle: true,
        languageDirs: Object.keys(flatTitlesMap[flatTitle]),
        titlesMap: flatTitlesMap
      };
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

    var titlesMap = {};

    languageDirs.forEach(function (langInfo) {
      var langDir = path.join(categoryRoot, langInfo.name);
      var files = fs.readdirSync(langDir, { withFileTypes: true })
        .filter(function (d) { return d.isFile() && path.extname(d.name).toLowerCase() === '.mov'; })
        .map(function (d) { return path.join(langDir, d.name); });

      files.forEach(function (fullPath) {
        var title = parseTitleFromFile(fullPath, langInfo.name);
        if (!title) return;
        if (!titlesMap[title]) titlesMap[title] = {};
        titlesMap[title][langInfo.canonicalName] = fullPath;
      });
    });

    return {
      isFlatSingleTitle: false,
      languageDirs: languageDirs.map(function (langInfo) { return langInfo.canonicalName; }),
      titlesMap: titlesMap
    };
  }

  function scanAllCategories(rootFolder) {
    var found = [];
    categoryOrder.forEach(function (categoryName) {
      var categoryPath = path.join(rootFolder, categoryName);
      if (fs.existsSync(categoryPath) && fs.statSync(categoryPath).isDirectory()) {
        found.push({
          name: categoryName,
          path: categoryPath,
          scanResult: scanSingleCategory(categoryName, categoryPath)
        });
      }
    });
    return found;
  }

  function chooseBatchForCategory(categoryName, scanResult, requestedCount, tracking) {
    tracking.categories = tracking.categories || {};
    tracking.categories[categoryName] = tracking.categories[categoryName] || { usedEnglishTitlesCycle: [], isFlatSingleTitle: !!scanResult.isFlatSingleTitle };
    tracking.usedLanguagesGlobalCycle = Array.isArray(tracking.usedLanguagesGlobalCycle)
      ? tracking.usedLanguagesGlobalCycle
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

    function getSaveTheEarthReuseMap(title, selectedLanguages) {
      var reuseMap = {};
      if (categoryName !== 'Save the Earth') return reuseMap;

      var selectedLookup = {};
      selectedLanguages.forEach(function (lang) { selectedLookup[lang] = true; });

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
      var bUnused = getUnusedLanguagesForTitle(b).length;
      var aUnused = getUnusedLanguagesForTitle(a).length;
      if (bUnused !== aUnused) return bUnused - aUnused;
      return getOtherLanguagesForTitle(b).length - getOtherLanguagesForTitle(a).length;
    });

    var chosenTitle = candidatePool[0];
    var languageToFile = getLanguageToFileForTitle(chosenTitle);
    var otherLanguages = Object.keys(languageToFile).filter(function (lang) { return lang !== 'English'; });
    var pickedOtherLanguages = shuffle(otherLanguages.filter(function (lang) {
      return tracking.usedLanguagesGlobalCycle.indexOf(lang) === -1;
    })).slice(0, otherNeeded);
    var warningMessages = [];
    var reusedFreshness = false;

    if (pickedOtherLanguages.length < otherNeeded && otherLanguages.length > pickedOtherLanguages.length) {
      var alreadyPicked = {};
      for (var i = 0; i < pickedOtherLanguages.length; i++) {
        alreadyPicked[pickedOtherLanguages[i]] = true;
      }
      var refillPool = otherLanguages.filter(function (lang) { return !alreadyPicked[lang]; });
      pickedOtherLanguages = pickedOtherLanguages.concat(shuffle(refillPool).slice(0, otherNeeded - pickedOtherLanguages.length));
      reusedFreshness = true;
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
      warningMessages.push(categoryName + ': not enough globally fresh non-English languages were available for the chosen title, so some languages were reused.');
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

  browseBtn.addEventListener('click', function () {
    try {
      if (window.cep && window.cep.fs && typeof window.cep.fs.showOpenDialogEx === 'function') {
        var result = window.cep.fs.showOpenDialogEx(false, true, 'Select the root folder that contains all 5 category folders');
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
  updateRepoInput.addEventListener('change', function () {
    persistSettings();
    updateState.latestRelease = null;
    updateState.latestVersion = '';
    setUpdateStatus(updateRepoInput.value.trim() ? 'Update repo saved. Click Check for Updates.' : 'Enter a GitHub repo like owner/repo to enable update checks.');
    setUpdateUiState();
  });
  checkUpdatesBtn.addEventListener('click', function () {
    checkForUpdates({ silent: false });
  });
  installUpdateBtn.addEventListener('click', installLatestUpdate);
  window.addEventListener('beforeunload', persistSettings);

  runBtn.addEventListener('click', function () {
    try {
      setStatus('Working...');
      if (!selectedRootFolder) {
        setStatus('Please choose the root folder that contains the 5 category folders first.');
        return;
      }

      var requestedCount = parseInt(slideCountInput.value, 10);
      var targetTrack = parseInt(targetTrackInput.value, 10);
      var ignoreV1 = !!ignoreV1Input.checked;

      if (!requestedCount || requestedCount < 1) {
        setStatus('Please enter a valid number of slides.');
        return;
      }
      if (!targetTrack || targetTrack < 1) {
        setStatus('Please enter a valid target video track number.');
        return;
      }

      persistSettings();
      var tracking = loadTracking();
      var categoryDataList = scanAllCategories(selectedRootFolder);
      if (!categoryDataList.length) {
        setStatus('No category folders were found under the selected root folder.');
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
          title: batch.chosenTitle,
          targetTrack: targetTrack,
          warning: batch.warning || ''
        });
        titleSummary.push(categoryData.name + ': ' + batch.chosenTitle);
        languageSummary.push(categoryData.name + ': ' + batch.selectedLanguages.join(', '));
        if (batch.warning) log('Warning: ' + batch.warning);
      });

      maybeResetGlobalLanguageCycle(tracking, categoryDataList);
      saveTracking(tracking);

      chosenTitleEl.textContent = titleSummary.join('\n');
      chosenLanguagesEl.textContent = languageSummary.join('\n');

      callJsx('newPeaceMakerImportAndPlaceMulti', {
        batches: batches,
        requestedCount: requestedCount,
        rootFolderName: path.basename(selectedRootFolder),
        targetTrack: targetTrack,
        ignoreV1: ignoreV1,
      }, function (result) {
        try {
          var parsed = JSON.parse(result);
          if (!parsed.ok) {
            setStatus('Error: ' + parsed.error);
            return;
          }
          var msg = 'Done.\n';
          for (var i = 0; i < parsed.results.length; i++) {
            var r = parsed.results[i];
            msg += r.categoryName + ' -> track V' + r.targetTrack + ': imported ' + r.importedCount + ', placed ' + r.placedCount; if (i === 0) { msg += ', shared interval ' + r.intervalSeconds + ' sec'; }
            if (r.warning) msg += ' | Warning: ' + r.warning;
            msg += '\n';
          }
          if (parsed.presetNote) msg += parsed.presetNote + '\n';
          if (parsed.note) msg += parsed.note + '\n';
          setStatus(msg);
        } catch (e) {
          setStatus('Raw response: ' + result);
        }
      });
    } catch (err) {
      setStatus('Error: ' + err.message);
    }
  });

  restoreSettings();
  updateState.installedVersion = readManifestVersion(manifestPath) || '';
  updateState.latestVersion = loadTracking().settings.lastAvailableVersion || '';
  setUpdateStatus(updateRepoInput.value.trim() ? 'Ready to check for updates.' : 'Enter a GitHub repo like owner/repo to enable update checks.');
  setUpdateUiState();
  if (shouldBackgroundCheck()) {
    checkForUpdates({ silent: true });
  }
})();
