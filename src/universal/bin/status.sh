#!/bin/bash

CHECK="com.codingkapoor.esslattendancereportgenerator.EsslAttendanceReportGenerator"
PID=`jps | grep EsslAttendanceReportGenerator | awk '{print $1}'`
STATUS=$(ps aux | grep -v grep | grep ${CHECK})

if [ "${#STATUS}" -gt 0 ] && [ -n ${PID} ]; then
    echo "`date`: EsslAttendanceReportGenerator is running"
else
    echo "`date`: EsslAttendanceReportGenerator is not running"
    exit 1
fi
