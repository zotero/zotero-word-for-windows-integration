modifiedFiles=`git status -s -u | wc -l`
if [ "$modifiedFiles" != "0" ]; then
	echo "dotm source is not up to date. Use unpack_dotm.sh to update source"
	exit 1
fi