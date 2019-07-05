package com.codingkapoor.esslattendancereportgenerator.model

import java.time.LocalDate

import com.codingkapoor.esslattendancereportgenerator.`package`.using
import com.codingkapoor.esslattendancereportgenerator.core.RuntimeEnvironment

import scala.io.Source

case class AttLog(empId: Int, date: LocalDate) extends Ordered[AttLog] {
  override def compare(that: AttLog): Int = {
    if (this.date == that.date) 0
    else if (this.date.isBefore(that.date)) -1
    else 1
  }
}

object AttLog {
  def getAttLogs: Set[AttLog] = {
    using(Source.fromFile(s"${RuntimeEnvironment.getDataDir}/1_attlog.dat")) { attlog =>
      attlog.getLines().toList.filter(l => l.trim.length > 0).map { line =>
        line.trim.split("\\s+") match {
          case Array(empId, date, _, _, _, _, _) =>
            AttLog(empId.toInt, LocalDate.parse(date))
        }
      }
    }.toSet
  }
}
