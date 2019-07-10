package com.codingkapoor.esslattendancereportgenerator.writer

import java.io.FileOutputStream
import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.`package`._
import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Employee, Holiday}
import com.codingkapoor.esslattendancereportgenerator.writer.attendance.AttendanceWriter
import com.codingkapoor.esslattendancereportgenerator.writer.attendanceheader.AttendanceHeaderWriter
import com.codingkapoor.esslattendancereportgenerator.writer.companydetails.CompanyDetailsWriter
import com.codingkapoor.esslattendancereportgenerator.writer.employeeinfo.EmployeeInfoWriter
import com.codingkapoor.esslattendancereportgenerator.writer.employeeinfoheader.EmployeeInfoHeaderWriter
import com.codingkapoor.esslattendancereportgenerator.writer.holiday.HolidayWriter
import com.codingkapoor.esslattendancereportgenerator.writer.sheetheader.SheetHeaderWriter
import com.codingkapoor.esslattendancereportgenerator.writer.weekend.WeekendWriter
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class ExcelWriter(val month: Int, val year: Int, val attendances: Seq[AttendancePerEmployee], val holidays: Seq[Holiday]) extends SheetHeaderWriter with CompanyDetailsWriter with
  EmployeeInfoHeaderWriter with EmployeeInfoWriter with AttendanceHeaderWriter with AttendanceWriter with
  HolidayWriter with WeekendWriter {
  val monthTitle: String = YearMonth.of(year, month).getMonth.toString

  private val employees: Seq[Employee] = attendances.map(l => l.employee)

  val attendanceDimensions: AttendanceDimensions = AttendanceDimensions(month, year, employees)
  val attendanceHeaderDimensions: AttendanceHeaderDimensions = AttendanceHeaderDimensions(month, year)
  val companyDetailsDimensions: CompanyDetailsDimensions= CompanyDetailsDimensions()
  val employeeInfoDimensions: EmployeeInfoDimensions = EmployeeInfoDimensions(employees)

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
    implicit val _attendances: Seq[AttendancePerEmployee] = attendances
    implicit val _holidays: Seq[Holiday] = holidays

    using(new XSSFWorkbook) { implicit workbook =>
      val yearMonth = YearMonth.of(year, month)
      val _month = yearMonth.getMonth.toString

      workbook.createSheet(_month)

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

object ExcelWriter {
  def apply(month: Int, year: Int, attendances: Seq[AttendancePerEmployee], holidays: Seq[Holiday]): ExcelWriter = {
    new ExcelWriter(month, year, attendances, holidays)
  }
}
