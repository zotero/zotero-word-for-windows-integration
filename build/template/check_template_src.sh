#!/bin/bash
cd "$( dirname "${BASH_SOURCE[0]}" )"
modifiedFiles=`git status -s -u | wc -l`
if [ "$modifiedFiles" != "0" ]; then
	echo "Template source is not up to date. Use unpack_templates.sh to update source"
	exit 1
fi
