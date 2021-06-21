#!/bin/sh

n=1

fio 4KSW.fio > /dev/null 2>&1 &

pid=$!
echo $pid

while [ $n -lt 5 ]
do
	sudo hdparm --fwdownload SYS_S.bin --yes-i-know-what-i-am-doing --please-destroy-my-drive /dev/sdb &
	pid2=$!
	echo $pid2
	wait $pid2
	sleep 10
	#kill -9 $pid2
done

