#!/bin/sh

f=1
n=1

#active="M190603J"
#inactive="M190604J"
active="A12345T"
inactive="A19905J"


fio 4KSW_1.fio --output=fio_$f.csv > /dev/null 2>&1 &
fioPID=$!
echo "fio pid: $fioPID"

fioStart=$(date "+Date: %Y.%m.%d, Time: %H:%M:%S")
fioS1=$(date +%s)
echo "... fio test $f started at: $fioStart ...">> FIO.log 2>&1

while [ $n -lt 201 ]
do
	echo "\nstep 1: fio"
	if pgrep -x fio >/dev/null
	then
		fioTime=$(date "+Date: %Y.%m.%d, Time: %H:%M:%S")
		echo "... fio test $f running ..."
		echo "... fio test $f running at: $fioTime ...">> FIO.log 2>&1
		
	else
		fioEnd=$(date "+Date: %Y.%m.%d, Time: %H:%M:%S")
		fioS2=$(date +%s)
		fioExe=$((fioS2-fioS1))
		
		echo "... fio test $f stopped ..."
		echo "... fio test $f finished at: $fioEnd, execution time: $fioExe sec ...">> FIO.log 2>&1
		echo "\n\n" >> FIO.log 2>&1
		
		f=$((f+1))
		
		fio 4KSW_1.fio --output=fio_$f.csv > /dev/null 2>&1 &
		fioPID=$!
		echo "fio pid: $fioPID"

		fioStart=$(date "+Date: %Y.%m.%d, Time: %H:%M:%S")
		fioS1=$(date +%s)
		echo "... fio test $f started at: $fioStart ...">> FIO.log 2>&1
	fi
	sleep 1
	
	echo "\nstep 2: hdparm"
	ffuStart=$(date "+Date: %Y.%m.%d, Time: %H:%M:%S")
	FFU_T1=$(date +%s)
	ffuS1=$(date +%s)
	
	echo "... hdparm is running ..."
    echo "FFU test $n started at: $ffuStart ..." >> FFU.log 2>&1
	sudo time hdparm --fwdownload $active.bin --yes-i-know-what-i-am-doing --please-destroy-my-drive /dev/sdb >> FFU.log 2>&1
	
	sata=$(sudo smartctl -a /dev/sdb | grep -i "SATA")
    fw=$(sudo smartctl -a /dev/sdb | grep -i "Firmware")
    echo "\n$fw\n$sata\n"
    
	ffuEnd=$(date "+Date: %Y.%m.%d, Time: %H:%M:%S")
	ffuS2=$(date +%s)
	ffuExe=$((ffuS2-ffuS1))
	sleep 1
	
	echo "step 3: status"
	N=3
	status=$(echo $fw | cut -d " " -f $N)
	echo "updated to: $status"
	
	if [ "$status" = "$active" ]; then
		echo "PASSED\n\n"
		echo "... ffu test $n finished at: $ffuEnd, execution time: $ffuExe sec ... " >> FFU.log 2>&1
		echo "... SATA: $sata\n...FW: $fw\n...status: PASSED\n\n\n\n\n" >> FFU.log 2>&1
	else
		echo "FAILED\n\n"
		echo "... ffu test $n finished at: $ffuEnd, execution time: $ffuExe sec ... " >> FFU.log 2>&1
		echo "... SATA: $sata\n...FW: $fw\n...status: FAILED\n\n\n\n\n" >> FFU.log 2>&1 
	fi

	date "+Date: %Y.%m.%d, Time: %H:%M:%S" >> DMESG.log 2>&1
	sudo dmesg | grep -E 'SATA|link|fail|error' >> DMESG.log 2>&1
    echo "\n\n\n\n\n" >> DMESG.log 2>&1
    
	sleep 1
	
	temp=$active
	active=$inactive
	inactive=$temp
	
	n=$((n+1))
	        
done
kill -9 $fioPID

echo "FFU Test Completed...I am happy now !!!"

