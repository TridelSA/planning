#!/bin/sh
# Usage : ./bump-version.sh  -> met à jour le numéro de version (?v=) dans index.html et sw.js d'un coup.
V=$(date +%Y%m%d%H%M)
sed -i -E "s/(style\.css\?v=)[0-9a-z]+/\1$V/g; s/(app\.js\?v=)[0-9a-z]+/\1$V/g" index.html sw.js
echo "Version : $V"
