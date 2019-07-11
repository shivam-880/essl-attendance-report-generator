package com.codingkapoor.esslattendancereportgenerator.model

import java.time.LocalDate

import com.codingkapoor.esslattendancereportgenerator.`package`._
import com.codingkapoor.esslattendancereportgenerator.core.RuntimeEnvironment

import scala.io.Source

case class AttendanceLog(empId: Int, date: LocalDate) extends Ordered[AttendanceLog] {
  override def compare(that: AttendanceLog): Int = {
    if (this.date == that.date) 0
    else if (this.date.isBefore(that.date)) -1
    else 1
  }
}

object AttendanceLog {
  def getAttendanceLogs(month: Int, year: Int): Map[Int, List[LocalDate]] = {
    using(Source.fromFile(s"${RuntimeEnvironment.getDataDir}/$AttendanceLogFileName")) { attlog =>
      attlog.getLines().toList.filter(l => l.trim.length > 0).map { line =>
        line.trim.split("\\s+") match {
          case Array(empId, date, _, _, _, _, _) =>
            AttendanceLog(empId.toInt, LocalDate.parse(date))
        }
      }
    }.toSet.foldLeft(Map[Int, List[LocalDate]]()) { (acc, att) =>
      if (att.date.getMonthValue == month && att.date.getYear == year) {
        if (acc.contains(att.empId)) {
          acc + (att.empId -> (att.date :: acc(att.empId)).sortWith(_.isBefore(_)))
        } else acc + (att.empId -> List(att.date).sortWith(_.isBefore(_)))
      } else acc
    }
  }
}
