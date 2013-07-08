#!/bin/sh

point="$1"

[ -z "$point" ] && echo "Usage: $0 point-version-number" && exit

cd "$(dirname "$0")"

prefix="' Common VBA Library, version "
version="`date +%Y-%m-%d`.$point"

sed -i "s/^$prefix.*$/$prefix$version/" VBALib_VERSION.bas
unix2dos VBALib_VERSION.bas
echo Updated version to $version
