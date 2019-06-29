package com.codingkapoor.esslattendancereportgenerator

import java.time.LocalDate

trait ConfigLoader extends RuntimeEnvironment {

  private val conf = appConf.getConfig("esslattendancereportgenerator")

  def getConfiguredMonthYear: (Int, Int) = {
    lazy val current = LocalDate.now

    if (conf.hasPath("month") && conf.hasPath("year")) {
      (conf.getInt("month"), conf.getInt("year"))
    } else (current.getMonthValue, current.getYear)
  }
}
