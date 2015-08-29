#!/bin/bash
rm -rf src/Zotero.dotm/*

unzip install/Zotero.dotm -d src/Zotero.dotm/

rm src/Zotero.dotm/word/vbaProject.bin
mkdir src/Zotero.dotm/word/vbaProject.bin
python tools/officeparser/officeparser.py -l ERROR -o src/Zotero.dotm/word/vbaProject.bin --extract-streams --extract-macros --extract-unknown-sectors install/Zotero.dotm
