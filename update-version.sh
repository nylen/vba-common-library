#!/bin/sh

point="$1"

[ -z "$point" ] && echo "Usage: $0 point-version-number" && exit

cd "$(dirname "$0")"

prefix="' Common VBA Library, version "
sed -i "s/^$prefix.*$/$prefix`date +%Y-%m-%d`.$point/" *.bas *.cls
unix2dos *.bas *.cls
