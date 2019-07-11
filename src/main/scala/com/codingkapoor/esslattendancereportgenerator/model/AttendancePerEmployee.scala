package com.codingkapoor.esslattendancereportgenerator.model

import java.time.{LocalDate, YearMonth}

import com.codingkapoor.esslattendancereportgenerator.AttendanceStatus
import com.codingkapoor.esslattendancereportgenerator.AttendanceStatus.AttendanceStatus

import scala.collection.mutable

case class AttendancePerEmployee(employee: Employee, attendance: Map[Int, AttendanceStatus])

object AttendancePerEmployee {
  def getAttendancePerEmployee(month: Int, year: Int, employee: Employee, att: List[LocalDate], holidays: Map[LocalDate, String], requests: Map[LocalDate, AttendanceStatus]): AttendancePerEmployee = {
    val yearMonth = YearMonth.of(year, month)
    val numOfDaysInMonth = yearMonth.lengthOfMonth

    val attendance = mutable.Map[Int, AttendanceStatus]()
    for (i <- 1 to numOfDaysInMonth) {
      val dateStr = s"$year-${"%02d".format(month)}-${"%02d".format(i)}"
      val date = LocalDate.parse(dateStr)

      val holiday = holidays.get(date)
      val request = requests.get(date)

      if (holiday.isDefined)
        attendance.put(i, AttendanceStatus.Holiday)
      else if (date.getDayOfWeek.toString.equals("SATURDAY"))
        attendance.put(i, AttendanceStatus.Saturday)
      else if (date.getDayOfWeek.toString.equals("SUNDAY"))
        attendance.put(i, AttendanceStatus.Sunday)
      else if (request.isDefined)
        attendance.put(i, request.get)
      else if (att.contains(date))
        attendance.put(i, AttendanceStatus.Present)
      else
        attendance.put(i, AttendanceStatus.Abscond)
    }

    AttendancePerEmployee(employee, attendance.toMap)
  }
}
