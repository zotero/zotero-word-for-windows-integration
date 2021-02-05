#!/bin/bash

# Zotero.dotm
echo 'Unpacking Zotero.dotm...'

cd "$( dirname "${BASH_SOURCE[0]}" )"
rm -rf Zotero.dotm/*

# Other than vbaProject.bin, all files are XML
unzip -q ../../install/Zotero.dotm -d Zotero.dotm/

# Pretty-print XML files
find Zotero.dotm/ -type f \( -iname '*.xml' -o -iname '*.rels' \) -exec  xmllint --output '{}' --format '{}' \;

# Extract vbaProject.bin
rm Zotero.dotm/word/vbaProject.bin
mkdir Zotero.dotm/word/vbaProject.bin
python ../tools/officeparser/officeparser.py -l ERROR -o Zotero.dotm/word/vbaProject.bin \
       --extract-macros ../../install/Zotero.dotm
