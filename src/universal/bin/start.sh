#!/bin/bash

BIN_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
HOME_DIR="$(dirname "$BIN_DIR")"
CONF_DIR="${HOME_DIR}/conf"
LIB_DIR="${HOME_DIR}/lib"
LOG_DIR="${HOME_DIR}/logs"
DATA_DIR="${HOME_DIR}/data"

if [ ! -d ${CONF_DIR} ] || [ ! -d ${LIB_DIR} ] || [ ! -d ${LOG_DIR} ]; then
  echo "`date`: Mandatory directory check failed."
  exit 0
fi

echo "Starting EsslAttendanceReportGenerator..."

nohup java -agentlib:jdwp=transport=dt_socket,server=y,suspend=n,address=5007 \
    -server -Dlogback.configurationFile="${CONF_DIR}/logback.xml" \
	-Dconf.dir="${CONF_DIR}" -Ddata.dir="${DATA_DIR}" -Dlogs.dir="${LOG_DIR}" -cp "${LIB_DIR}/*:${CONF_DIR}/*" \
	com.codingkapoor.esslattendancereportgenerator.EsslAttendanceReportGenerator > "${LOG_DIR}/stdout.log" 2>&1 &

essl_attendance_report_generator_pid=$!

echo
echo "EsslAttendanceReportGenerator started with PID [$essl_attendance_report_generator_pid] at [`date`]"
