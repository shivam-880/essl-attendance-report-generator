package com.codingkapoor.esslattendancereportgenerator.core

import com.codingkapoor.esslattendancereportgenerator.writer._
import com.typesafe.scalalogging.LazyLogging

object EsslAttendanceReportGenerator extends App with LazyLogging with ConfigLoader {
  val (month, year) = getConfiguredMonthYear
  logger.debug(s"month = $month, year = $year")

  ExcelWriter(month, year).write()
}
