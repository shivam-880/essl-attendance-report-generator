package com.codingkapoor.esslattendancereportgenerator

import java.time.LocalDate

import scala.io.Source

case class Attendance(empId: Int, date: LocalDate) extends Ordered[Attendance] {
  override def compare(that: Attendance): Int = {
    if (this.date == that.date) 0
    else if (this.date.isBefore(that.date)) -1
    else 1
  }
}

object Attendance {
  def getAttendance: Set[Attendance] = {
    using(Source.fromFile(s"${RuntimeEnvironment.getDataDir}/1_attlog.dat")) { attlog =>
      attlog.getLines().toList.filter(l => l.trim.length > 0).map { line =>
        line.trim.split("\\s+") match {
          case Array(empId, date, _, _, _, _, _) =>
            Attendance(empId.toInt, LocalDate.parse(date))
        }
      }
    }.toSet
  }
}
