'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Suppress warnings about deprecated APIs
build.addSuppression(/Warning - \[sass\]/);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

// Configure TypeScript
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    // Add any custom webpack configuration here
    return generatedConfiguration;
  }
});

// Disable tslint (use eslint instead)
build.tslintCmd.enabled = false;

// Initialize build
build.initialize(require('gulp'));
