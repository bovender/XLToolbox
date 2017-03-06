#!/bin/sh
if [ -z $1 ]; then
        echo "Usage: $0 version"
        echo "where version is the bovender version (major.minor.patch)"
        exit 1
fi
find XLToolbox XLToolboxForExcel Tests -name '*.csproj' -print0 -o -name 'packages.config' -print0 | \
xargs -0 sed -ri 's/(Bovender[^0-9]+)([0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2})/\1'"$1/"
