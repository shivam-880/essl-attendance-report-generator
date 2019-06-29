package com.codingkapoor.esslattendancereportgenerator

import com.typesafe.scalalogging.LazyLogging

object EsslAttendanceReportGenerator extends App with LazyLogging with RuntimeEnvironment {
  logger.info(s"Configured interface = ${appConf.getConfig("esslattendancereportgenerator").getString("interface")}")
}
