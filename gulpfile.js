'use strict';

const build = require('@microsoft/sp-build-web');
const path = require('path');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

// ── @pnp/spfx-controls-react webpack alias ──────────────────────────────────
// The package uses a bare 'ControlStrings' import that resolves to its own
// localisation file. Without this alias webpack cannot bundle the control.
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    if (!generatedConfiguration.resolve) {
      generatedConfiguration.resolve = {};
    }
    if (!generatedConfiguration.resolve.alias) {
      generatedConfiguration.resolve.alias = {};
    }
    generatedConfiguration.resolve.alias['ControlStrings'] = path.resolve(
      __dirname,
      'node_modules/@pnp/spfx-controls-react/lib/loc/en-us.js'
    );
    return generatedConfiguration;
  }
});

build.initialize(require('gulp'));
