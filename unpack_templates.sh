#!/bin/bash

# Zotero.dotm
echo 'Unpacking Zotero.dotm...'

rm -rf src/Zotero.dotm/*

# Other than vbaProject.bin, all files are XML
unzip -q install/Zotero.dotm -d src/Zotero.dotm/

# Pretty-print XML files
find src/Zotero.dotm/ -type f \( -iname '*.xml' -o -iname '*.rels' \) -exec  xmllint --output '{}' --format '{}' \;

# Extract vbaProject.bin
rm src/Zotero.dotm/word/vbaProject.bin
mkdir src/Zotero.dotm/word/vbaProject.bin
python tools/officeparser/officeparser.py -l ERROR -o src/Zotero.dotm/word/vbaProject.bin --extract-streams --extract-macros --extract-unknown-sectors install/Zotero.dotm

# Zotero.dot (not unzipping, because all files are binary anyway)
echo 'Unpacking Zotero.dot...'

rm -rf src/Zotero.dot/*

python tools/officeparser/officeparser.py -l ERROR -o src/Zotero.dot --extract-streams --extract-macros --extract-unknown-sectors install/Zotero.dot