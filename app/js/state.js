(function () {
  var state = {
    activeModule: 'home',
    activeProject: null,
    savedProjects: [],
    storm: {},
    condensate: {},
    vent: {},
    gas: {},
    wsfu: {},
    fixtureUnit: {},
    solar: {},
    duct: {},
    ductStatic: {},
    refrigerant: {}
  };

  window.appState = state;

  window.updateAppState = function updateAppState(path, value) {
    if (!path) return;
    var parts = String(path).split('.');
    var node = state;
    for (var i = 0; i < parts.length - 1; i++) {
      var key = parts[i];
      if (!Object.prototype.hasOwnProperty.call(node, key) || typeof node[key] !== 'object' || node[key] === null) {
        node[key] = {};
      }
      node = node[key];
    }
    node[parts[parts.length - 1]] = value;
  };
})();
