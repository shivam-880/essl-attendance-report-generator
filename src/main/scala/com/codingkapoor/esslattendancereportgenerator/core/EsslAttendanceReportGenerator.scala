package com.codingkapoor.esslattendancereportgenerator.core

import java.time.LocalDate

import com.codingkapoor.esslattendancereportgenerator.model.{AttLog, Employee, Holiday, Request}
import com.typesafe.scalalogging.LazyLogging

// App you reply with an error message if uptodate attlog or holidays is not provided
// Also notify entries in attlog that were not mentioned in employees.json
object EsslAttendanceReportGenerator extends App with LazyLogging with ConfigLoader {
  val (month, year) = getConfiguredMonthYear
  logger.debug(s"month = $month, year = $year")

  val holidays = Holiday.getHolidays.filter(h => h.date.getMonthValue == month && h.date.getYear == year)
  logger.debug(s"holidays = $holidays")

  val employees = Employee.getEmployees
  logger.debug(s"employees = $employees")

  val attLogs = AttLog.getAttLogs.foldLeft(Map[Int, List[LocalDate]]()) { (acc, att) =>
    if (acc.contains(att.empId)) {
      acc + (att.empId -> (att.date :: acc(att.empId)).sortWith(_.isBefore(_)))
    } else acc + (att.empId -> List(att.date).sortWith(_.isBefore(_)))
  }
  logger.debug(s"attLogs = $attLogs")

  val requests = Request.getRequests
  logger.debug(s"requests = $requests")

  ExcelWriter.write(month, year, employees)
}
