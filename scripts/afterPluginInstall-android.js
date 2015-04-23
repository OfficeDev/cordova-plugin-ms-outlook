#!/usr/bin/env node

module.exports = function (ctx) {
    var path = ctx.requireCordovaModule('path');
    var et = ctx.requireCordovaModule('elementtree');
    var ConfigParser = ctx.requireCordovaModule('../ConfigParser/ConfigParser');
    var configFile = path.resolve(ctx.opts.projectRoot, 'config.xml');
    var config = new ConfigParser(configFile);
    var MIN_SDK_VERSION = 15;
    var PREF_NAME = 'android-minSdkVersion';

    var minSdkVersion = config.getPreference(PREF_NAME);

    function setGlobalPreference(config, name, value) {
        var pref = config.doc.find('preference[@name="' + name + '"]');
        if (!pref) {
            pref = new et.Element('preference');
            pref.attrib.name = name;
            config.doc.getroot().append(pref);
        }
        pref.attrib.value = value;

        config.write();
    }

    // Add required minSdkVersion to config or change min version if it was less than we need
    // TODO: Replace with this when Cordova tools including setGlobalPreference is released
    // config.setGlobalPreference(PREF_NAME, MIN_SDK_VERSION);
    if(!minSdkVersion || (parseInt(minSdkVersion, 10) < MIN_SDK_VERSION)) {
        setGlobalPreference(config, PREF_NAME, MIN_SDK_VERSION);
    }
};
