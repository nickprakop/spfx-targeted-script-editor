'use strict';

const build = require('@microsoft/sp-build-web');
const gulp = require('gulp');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

const crntConfig = build.getConfig();


// Extend the SPFx build rig, and overwrite the `shouldWarningsFailBuild` property
if (true) {
  class CustomSPWebBuildRig extends build.SPWebBuildRig {
    setupSharedConfig() {
      build.log("IMPORTANT: Warnings will not fail the build.")
      build.mergeConfig({
        shouldWarningsFailBuild: false
      });
      super.setupSharedConfig();
    }
  }

  build.rig = new CustomSPWebBuildRig();
}

/* fast-serve */
const { addFastServe } = require("spfx-fast-serve-helpers");
addFastServe(build);
/* end of fast-serve */

build.initialize(gulp);

