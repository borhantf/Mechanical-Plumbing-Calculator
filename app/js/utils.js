window.Utils = window.Utils || {
  roundTo: function roundTo(value, digits) {
    var d = (typeof digits === 'number' ? digits : 2);
    var n = Number(value);
    if (!isFinite(n)) return 0;
    var f = Math.pow(10, d);
    return Math.round(n * f) / f;
  },
  isPositiveNumber: function isPositiveNumber(value) {
    var n = Number(value);
    return isFinite(n) && n > 0;
  },
  isNonNegativeNumber: function isNonNegativeNumber(value) {
    var n = Number(value);
    return isFinite(n) && n >= 0;
  },
  setWarningHtml: function setWarningHtml(node, message) {
    if (!node) return;
    node.innerHTML = '<div class="warning-line">' + String(message || '') + '</div>';
  },
  setPlaceholderHtml: function setPlaceholderHtml(node, message) {
    if (!node) return;
    node.innerHTML = '<div class="result-placeholder">' + String(message || '') + '</div>';
  }
};
