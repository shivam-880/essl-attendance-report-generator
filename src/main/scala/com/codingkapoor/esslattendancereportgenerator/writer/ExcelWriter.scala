package com.codingkapoor.esslattendancereportgenerator.writer

import java.io.FileOutputStream
import java.time.{LocalDate, YearMonth}

import com.codingkapoor.esslattendancereportgenerator.`package`._
import com.codingkapoor.esslattendancereportgenerator.model._
import com.codingkapoor.esslattendancereportgenerator.writer.attendance.AttendanceWriter
import com.codingkapoor.esslattendancereportgenerator.writer.attendanceheader.AttendanceHeaderWriter
import com.codingkapoor.esslattendancereportgenerator.writer.companydetails.CompanyDetailsWriter
import com.codingkapoor.esslattendancereportgenerator.writer.employeeinfo.EmployeeInfoWriter
import com.codingkapoor.esslattendancereportgenerator.writer.employeeinfoheader.EmployeeInfoHeaderWriter
import com.codingkapoor.esslattendancereportgenerator.writer.holiday.HolidayWriter
import com.codingkapoor.esslattendancereportgenerator.writer.sheetheader.SheetHeaderWriter
import com.codingkapoor.esslattendancereportgenerator.writer.weekend.WeekendWriter
import com.typesafe.scalalogging.LazyLogging
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import scala.collection.immutable

class ExcelWriter private(val month: Int, val year: Int, val attendances: Seq[AttendancePerEmployee], val employees: Seq[Employee], val holidays: Seq[Holiday]) extends SheetHeaderWriter with CompanyDetailsWriter with
  EmployeeInfoHeaderWriter with EmployeeInfoWriter with AttendanceHeaderWriter with AttendanceWriter with
  HolidayWriter with WeekendWriter {
  val yearMonth: YearMonth = YearMonth.of(year, month)
  val monthTitle: String = yearMonth.getMonth.toString

  val attendanceDimensions: AttendanceDimensions = AttendanceDimensions(month, year, employees)
  val attendanceHeaderDimensions: AttendanceHeaderDimensions = AttendanceHeaderDimensions(month, year)
  val companyDetailsDimensions: CompanyDetailsDimensions = CompanyDetailsDimensions()
  val employeeInfoDimensions: EmployeeInfoDimensions = EmployeeInfoDimensions(employees)
  val employeeInfoHeaderDimensions: EmployeeInfoHeaderDimensions = EmployeeInfoHeaderDimensions()
  val sectionHeaderDimensions: SectionHeaderDimensions = SectionHeaderDimensions(month, year)

  override def mergedRegionAlreadyExists(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int)(implicit workbook: XSSFWorkbook): Boolean = {
    val sheet = workbook.getSheet(monthTitle)

    var existingMergedRegionsFound = List.empty[Boolean]
    val mergedRegions = sheet.getMergedRegions.iterator()
    while (mergedRegions.hasNext) {
      val cellRangeAddr = mergedRegions.next()
      if (cellRangeAddr.getFirstRow == firstRowIndex && cellRangeAddr.getLastRow == lastRowIndex &&
        cellRangeAddr.getFirstColumn == firstColumnIndex && cellRangeAddr.getLastColumn == lastColumnIndex)
        existingMergedRegionsFound = true :: existingMergedRegionsFound
    }

    existingMergedRegionsFound.contains(true)
  }

  def write(): Unit = {
    using(new XSSFWorkbook) { implicit workbook =>
      workbook.createSheet(monthTitle)

      writeSheetHeader
      writeCompanyDetails
      writeEmployeeInfoHeader
      writeEmployeeInfo
      writeAttendanceHeader
      writeAttendance
      writeHolidays
      writeWeekends

      using(new FileOutputStream(AttendanceReportFileName))(workbook.write)
    }
  }
}

object ExcelWriter extends LazyLogging {
  def apply(month: Int, year: Int): ExcelWriter = {
    val holidays = Holiday.getHolidays.filter(h => h.date.getMonthValue == month && h.date.getYear == year)
    logger.debug(s"holidays = $holidays")

    val employees = Employee.getEmployees
    logger.debug(s"employees = $employees")

    val attendanceLogs: Map[Int, List[LocalDate]] = AttendanceLog.getAttendanceLogs.foldLeft(Map[Int, List[LocalDate]]()) { (acc, att) =>
      if (acc.contains(att.empId)) {
        acc + (att.empId -> (att.date :: acc(att.empId)).sortWith(_.isBefore(_)))
      } else acc + (att.empId -> List(att.date).sortWith(_.isBefore(_)))
    }
    logger.debug(s"attendanceLogs = $attendanceLogs")

    val requests = Request.getRequests
    logger.debug(s"requests = $requests")

    val attendances: immutable.Seq[AttendancePerEmployee] = employees.map { employee =>
      val attendance = attendanceLogs(employee.empId) // handle key not found
    val _requests: List[Request] = requests.filter(l => l.empId == employee.empId)
      val r = AttendancePerEmployee.getAttendancePerEmployee(employee, attendance, holidays, _requests)(month, year)
      logger.debug(s"emp = ${employee.empId}, r = ${r.attendance.toList.sortWith(_._1 < _._1)}")

      r
    }

    new ExcelWriter(month, year, attendances, employees, holidays)
  }
}
