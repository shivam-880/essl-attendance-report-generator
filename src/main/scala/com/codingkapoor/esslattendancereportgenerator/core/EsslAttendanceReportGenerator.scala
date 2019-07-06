package com.codingkapoor.esslattendancereportgenerator.core

import java.time.LocalDate

import com.codingkapoor.esslattendancereportgenerator.model._
import com.typesafe.scalalogging.LazyLogging

import scala.collection.immutable

// App you reply with an error message if uptodate attlog or holidays is not provided
// Also notify entries in attlog that were not mentioned in employees.json
object EsslAttendanceReportGenerator extends App with LazyLogging with ConfigLoader {
  val (month, year) = getConfiguredMonthYear
  logger.debug(s"month = $month, year = $year")

  val holidays = Holiday.getHolidays.filter(h => h.date.getMonthValue == month && h.date.getYear == year)
  logger.debug(s"holidays = $holidays")

  val employees = Employee.getEmployees
  logger.debug(s"employees = $employees")

  val attendanceLogs: Map[Int, List[LocalDate]] = AttendanceLog.getAttendanceLogs.foldLeft(Map[Int, List[LocalDate]]()) { (acc, att) =>
    if (acc.contains(att.empId)) {
      acc + (att.empId -> (att.date :: acc(att.empId)).sortWith(_.isBefore(_)))
    } else acc + (att.empId -> List(att.date).sortWith(_.isBefore(_)))
  }
  logger.debug(s"attendanceLogs = $attendanceLogs")

  val requests = Request.getRequests
  logger.debug(s"requests = $requests")

  val attendances: immutable.Seq[AttendancePerEmployee] = employees.map { employee =>
    val r = AttendancePerEmployee.getAttendancePerEmployee(employee, attendanceLogs(employee.empId), holidays, requests.filter(l => l.empId == employee.empId))(month, year)
    logger.debug(s"emp = ${employee.empId}, r = ${r.attendance.toList.sortWith(_._1 < _._1)}")

    r
  }

  ExcelWriter.write(attendances)(month, year)
}
