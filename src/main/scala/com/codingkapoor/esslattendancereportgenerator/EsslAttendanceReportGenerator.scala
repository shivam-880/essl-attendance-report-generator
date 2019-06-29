package com.codingkapoor.esslattendancereportgenerator

import com.typesafe.scalalogging.LazyLogging

object EsslAttendanceReportGenerator extends App with LazyLogging with ConfigLoader {
  val (month, year) = getConfiguredMonthYear
  logger.info(s"Configured month = $month, year = $year")
}
