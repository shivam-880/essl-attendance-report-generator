package com.codingkapoor.esslattendancereportgenerator

import com.typesafe.scalalogging.LazyLogging

// App you reply with an error message if uptodate attlog or holidays is not provided
// Also notify entries in attlog that were not mentioned in employees.json
object EsslAttendanceReportGenerator extends App with LazyLogging with ConfigLoader {
  val (month, year) = getConfiguredMonthYear
  logger.info(s"Configured month = $month, year = $year")
  logger.info(s"Holidays = ${Holiday.getHolidays.filter(h => h.date.getMonthValue == month && h.date.getYear == year)}")
  logger.info(s"Employees = ${Employee.getEmployees}")
  logger.info(s"Attendance = ${Attendance.getAttendance}")
  logger.info(s"Requests = ${Request.getRequests}")
}
