package com.codingkapoor.esslattendancereportgenerator

import java.time.LocalDate

import scala.io.Source

case class Attendance(empId: Int, date: LocalDate)

object Attendance {
  def getAttendance: List[Attendance] = {
    using(Source.fromFile(s"${RuntimeEnvironment.getDataDir}/1_attlog.dat")) { attlog =>
      attlog.getLines().toList.filter(l => l.trim.length > 0).map { line =>
        line.trim.split("\\s+") match {
          case Array(empId, date, _, _, _, _, _) =>
            Attendance(empId.toInt, LocalDate.parse(date))
        }
      }
    }
  }
}
