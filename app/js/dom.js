window.DOM = window.DOM || {
  byId: function byId(id) { return document.getElementById(id); },
  query: function query(selector, root) { return (root || document).querySelector(selector); },
  queryAll: function queryAll(selector, root) { return Array.prototype.slice.call((root || document).querySelectorAll(selector)); }
};
