package com.codingkapoor.esslattendancereportgenerator.model

import java.time.{LocalDate, YearMonth}

import com.codingkapoor.esslattendancereportgenerator.AttendanceStatus
import com.codingkapoor.esslattendancereportgenerator.`package`.Attendance

import scala.collection.mutable

case class AttendancePerEmployee(employee: Employee, attendance: Attendance)

object AttendancePerEmployee {
  def getAttendancePerEmployee(employee: Employee, att: List[LocalDate], holidays: Seq[Holiday], requests: Seq[Request])(month: Int, year: Int): AttendancePerEmployee = {
    val yearMonth = YearMonth.of(year, month)
    val numOfDays = yearMonth.lengthOfMonth

    val holidaysMap: Map[LocalDate, String] = holidays.map(h => h.date -> h.occasion).toMap
    val requestsMap: Map[LocalDate, String] = requests.map(r => r.date -> r.req).toMap

    val attendance = mutable.Map[Int, String]()
    for(i <- 1 to numOfDays) {
      val dateStr = s"$year-${"%02d".format(month)}-${"%02d".format(i)}"
      val date = LocalDate.parse(dateStr)
      val holiday = holidaysMap.get(date)
      val request = requestsMap.get(date)

      if(holiday.isDefined) {
        attendance.put(i, AttendanceStatus.Holiday.toString)
      } else if (date.getDayOfWeek.toString.equals("SATURDAY")) {
        attendance.put(i, AttendanceStatus.Saturday.toString)
      } else if (date.getDayOfWeek.toString.equals("SUNDAY")) {
        attendance.put(i, AttendanceStatus.Sunday.toString)
      } else if(request.isDefined) {
        attendance.put(i, request.get)
      } else if(att.contains(date)) {
        attendance.put(i, AttendanceStatus.Present.toString)
      } else {
        attendance.put(i, AttendanceStatus.Abscond.toString)
      }
    }

    AttendancePerEmployee(employee, attendance)
  }
}
