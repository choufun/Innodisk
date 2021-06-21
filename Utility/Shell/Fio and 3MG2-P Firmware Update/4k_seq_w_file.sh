
dev='test.file'




fio --rw=rw --size=1G --rwmixread=50 --name=test_rand4k --bs=4k --direct=1 --filename=${dev} --numjobs=1 --ioengine=libaio --runtime=5m --iodepth=1 --norandommap --time_based -write_bw_log=4KRanRead  --log_avg_msec=1000

	
