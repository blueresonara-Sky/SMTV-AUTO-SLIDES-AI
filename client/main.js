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
  var installedVersionEl = document.getElementById('installedVersion');
  var latestVersionEl = document.getElementById('latestVersion');
  var installUpdateBtn = document.getElementById('installUpdateBtn');
  var updateStatusEl = document.getElementById('updateStatus');
  var statusEl = document.getElementById('status');
  var chosenTitleEl = document.getElementById('chosenTitle');
  var chosenLanguagesEl = document.getElementById('chosenLanguages');

  var selectedRootFolder = '';
  var extensionRoot = '';
  var manifestPath = '';
  var trackingDir = path.join(os.homedir(), '.new-peace-maker');
  var trackingFile = path.join(trackingDir, 'usage-history.json');
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
    tracking.settings.slideAnchor = slideAnchorInput ? String(slideAnchorInput.value || 'top-right') : 'top-right';
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
  }

  function setUpdateStatus(msg) {
    updateStatusEl.textContent = msg;
  }

  function setUpdateUiState() {
    var hasUpdate = !!updateState.latestRelease && compareVersions(updateState.latestVersion, updateState.installedVersion) > 0;
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

  function getGitHubReleaseApiUrl(repo) {
    return 'https://api.github.com/repos/' + repo + '/releases/latest';
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
      '$Version = ' + quotePowerShellLiteral(latestVersion || ''),
      '',
      'function Wait-ForPremiereExit {',
      '  $deadline = (Get-Date).AddMinutes(30)',
      '  while ((Get-Date) -lt $deadline) {',
      "    $running = Get-Process -Name 'Adobe Premiere Pro','CEPHtmlEngine' -ErrorAction SilentlyContinue",
      '    if (-not $running) { return }',
      '    Start-Sleep -Seconds 2',
      '  }',
      '}',
      '',
      'function Install-UpdateFiles {',
      '  if (-not (Test-Path -LiteralPath $TargetDir)) {',
      '    New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null',
      '  }',
      '',
      '  for ($i = 0; $i -lt 120; $i++) {',
      '    try {',
      '      Get-ChildItem -LiteralPath $TargetDir -Force -ErrorAction SilentlyContinue | ForEach-Object {',
      '        Remove-Item -LiteralPath $_.FullName -Recurse -Force',
      '      }',
      '      Get-ChildItem -LiteralPath $SourceDir -Force | ForEach-Object {',
      '        Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $TargetDir $_.Name) -Recurse -Force',
      '      }',
      '      return',
      '    } catch {',
      '      Start-Sleep -Seconds 2',
      '    }',
      '  }',
      "  throw 'Could not replace the extension files after waiting for Premiere Pro to close.'",
      '}',
      '',
      'try {',
      '  Wait-ForPremiereExit',
      '  Install-UpdateFiles',
      '  try { Remove-Item -LiteralPath $TempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}',
      '} catch {',
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

    requestJson(getGitHubReleaseApiUrl(updateRepo), function (err, release) {
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

    if (!window.confirm('Install version ' + latestVersion + ' from ' + updateRepo + ' now? Premiere Pro should be restarted after the update.')) {
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

        if (process.platform === 'win32') {
          setUpdateStatus('Preparing update installer...');
          launchWindowsDeferredInstaller(extractedExtensionRoot, tempRoot, latestVersion);
          updateState.installing = false;
          setUpdateStatus('Update is staged. Accept the Windows prompt, then close Premiere Pro. The installer will finish after Premiere exits and update to version ' + latestVersion + '.');
          setUpdateUiState();
          return;
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

      callJsx('newPeaceMakerImportAndPlaceMulti', {
        batches: batches,
        requestedCount: requestedCount,
        rootFolderName: path.basename(selectedRootFolder),
        targetTrack: targetTrack,
        ignoreV1: ignoreV1,
        slideAnchor: slideAnchor,
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

  extensionRoot = resolveExtensionRoot();
  manifestPath = extensionRoot ? path.join(extensionRoot, 'CSXS', 'manifest.xml') : path.join(path.resolve(__dirname, '..'), 'CSXS', 'manifest.xml');
  restoreSettings();
  updateState.installedVersion = readManifestVersion(manifestPath) || '';
  updateState.latestVersion = loadTracking().settings.lastAvailableVersion || '';
  setUpdateStatus(updateRepo ? 'Ready to check for updates.' : 'No GitHub update source is configured.');
  setUpdateUiState();
  if (updateRepo) {
    checkForUpdates({ silent: true });
  }
})();
