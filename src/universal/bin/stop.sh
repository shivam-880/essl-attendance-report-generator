#!/bin/bash

PID=`jps | grep EsslAttendanceReportGenerator | awk '{print $1}'`

COUNTER=0

if [ "${#PID}" -gt 0 ]
then
	kill ${PID}
	while [[ ( -d /proc/${PID} ) && ( -z `grep zombie /proc/${PID}/status` ) ]]
	do
    	echo "`date`: Waiting for EsslAttendanceReportGenerator to gracefully stop"
    	sleep 10
    	COUNTER=$(($COUNTER+1))
    	if [ "$COUNTER" -gt 10 ]; then kill -9 ${PID}; break; fi
	done
	echo "`date`: EsslAttendanceReportGenerator stopped"
else
	echo "`date`: EsslAttendanceReportGenerator was not running"
fi
