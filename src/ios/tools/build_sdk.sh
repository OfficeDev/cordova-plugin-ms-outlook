#!/bin/bash

# Copyright (c) Microsoft Open Technologies, Inc.
# All Rights Reserved
# Licensed under the Apache License, Version 2.0.
# See License.txt in the project root for license information.

# Implements logic to build https://github.com/OfficeDev/Office-365-SDK-for-iOS and produce required libs
# Usage: place this script to sdk-objectivec folder and run

BUILD_PATH="build"
BUILD_CONFIGURATION="Debug"

PROJECTS_TO_BUILD=(office365_odata_base office365_exchange_sdk)

for i in "${PROJECTS_TO_BUILD[@]}"
do
	proj="${i}"
	echo "Building $proj" 
	xcodebuild -workspace office365-services.xcworkspace -scheme $proj -configuration $BUILD_CONFIGURATION ARCHS="i386 x86_64" -sdk iphonesimulator VALID_ARCHS="i386 x86_64" ONLY_ACTIVE_ARCH=NO CONFIGURATION_BUILD_DIR="../build/emulator" clean build
	xcodebuild -workspace office365-services.xcworkspace -scheme $proj -configuration $BUILD_CONFIGURATION ARCHS="armv7 armv7s arm64" -sdk iphoneos VALID_ARCHS="armv7 armv7s arm64" CONFIGURATION_BUILD_DIR="../build/device" clean build
	echo "Creating universal version of $proj"
	rm -rf "$BUILD_PATH/$proj.framework"
	cp -R "$BUILD_PATH/emulator/$proj.framework" "$BUILD_PATH/$proj.framework"

	simulatorLibPath="$BUILD_PATH/emulator/$proj.framework/$proj"
	deviceLibPath="$BUILD_PATH/device/$proj.framework/$proj"
	universalLibPath="$BUILD_PATH/$proj.framework/$proj"

	lipo "$simulatorLibPath" "$deviceLibPath" -create -output "$universalLibPath"
	lipo -info "$universalLibPath"
done

echo "Done. Build artifacts could be found at '$BUILD_PATH' folder"