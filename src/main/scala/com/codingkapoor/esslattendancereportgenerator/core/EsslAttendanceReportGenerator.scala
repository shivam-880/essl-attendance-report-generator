package com.codingkapoor.esslattendancereportgenerator.core

import com.codingkapoor.esslattendancereportgenerator.writer._
import com.typesafe.scalalogging.LazyLogging

// App you reply with an error message if uptodate attlog or holidays is not provided
// Also notify entries in attlog that were not mentioned in employees.json
object EsslAttendanceReportGenerator extends App with LazyLogging with ConfigLoader {
//  Thread.sleep(2000)
  val (month, year) = getConfiguredMonthYear
  logger.debug(s"month = $month, year = $year")

  ExcelWriter(month, year).write()
}
