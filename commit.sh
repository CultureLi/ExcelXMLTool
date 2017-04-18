git add *.xml

args=""
remote=`ps -ef|git status`
OLD_IFS="$IFS"
IFS="/"
array=($remote)
resultMessage="success"

branch=''
cd $PWD
if [ -d '.git' ]; then
    output=`git describe --contains --all HEAD|tr -s '\n'`
    if [ "$output" ]; then
        branch="${output}"
    fi
fi

git fetch origin $branch


Current=$(git rev-parse HEAD)
Origin=$(git rev-parse FETCH_HEAD)
if [ "$Current" != "$Origin" ];then
	echo "请先Pull至最新节点"
	exit 1
fi


for((i=0;i<${#array[@]};i++))
do
	if [[ "${array[i]}" == *.xml* ]]
	then
		path=${array[i]}
		#取出文件名
		OLD_IFS="$IFS"
		IFS=" "
		pathArray=($path)
		for((j=0;j<${#pathArray[@]};j++))
		do
			if [[ "${pathArray[j]}" == *.xml* ]]
			then
				fileName="${pathArray[j]}"
				echo $fileName
				OLD_IFS="$IFS"
				IFS="."
				fileNameArray=($fileName)
				IFS=""
				resultMessage=`ps -ef|./ExcelXmlTool.exe "${fileNameArray[0]}".xml`
				if [[ "$resultMessage" == *error* ]]
				then
					break
				fi
			fi
		done
	fi
done

if [[ "$resultMessage" == "success" ]]; then
	TortoiseGitProc.exe /command:commit
else
	echo $resultMessage
	read -n 1
fi


