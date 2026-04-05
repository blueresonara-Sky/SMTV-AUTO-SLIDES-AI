/* Minimal CSInterface shim placeholder. Replace with Adobe official CSInterface.js if needed. */
(function () {
  if (window.CSInterface) return;
  function CSInterface() {}
  CSInterface.prototype.evalScript = function (script, callback) {
    try {
      if (typeof window.__adobe_cep__ !== 'undefined' && window.__adobe_cep__.evalScript) {
        window.__adobe_cep__.evalScript(script, callback || function () {});
      } else if (callback) {
        callback('');
      }
    } catch (e) {
      if (callback) callback(String(e));
    }
  };
  window.CSInterface = CSInterface;
})();
