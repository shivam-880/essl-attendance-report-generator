package com.codingkapoor.esslattendancereportgenerator.writer

import java.io.FileOutputStream
import java.sql.Date
import java.time.{DayOfWeek, LocalDate, YearMonth}

import com.codingkapoor.esslattendancereportgenerator.AttendanceStatus
import com.codingkapoor.esslattendancereportgenerator.`package`._
import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Holiday}
import org.apache.poi.ss.usermodel._
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder
import org.apache.poi.xssf.usermodel.{XSSFColor, XSSFWorkbook}

class ExcelWriter(val month: Int, val year: Int) extends
  SheetHeaderWriter with CompanyDetailsWriter with AttendanceHeaderWriter with AttendanceWriter with HolidayWriter with WeekendWriter {

  override def mergedRegionAlreadyExists(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int)(implicit workbook: XSSFWorkbook): Boolean = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString

    val sheet = workbook.getSheet(_month)

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

  private def resizeCols(implicit workbook: XSSFWorkbook): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString
    val numOfDaysInMonth = yearMonth.lengthOfMonth

    val sheet = workbook.getSheet(_month)

    for (i <- EmployeesInfoHeader.indices)
      sheet.autoSizeColumn(i)

    for (i <- 5 to (5 + numOfDaysInMonth))
      sheet.setColumnWidth(i, 1200)
  }

  def write(attendances: Seq[AttendancePerEmployee], holidays: Seq[Holiday]): Unit = {
    implicit val _attendances: Seq[AttendancePerEmployee] = attendances
    implicit val _holidays: Seq[Holiday] = holidays

    using(new XSSFWorkbook) { implicit workbook =>
      val yearMonth = YearMonth.of(year, month)
      val _month = yearMonth.getMonth.toString

      workbook.createSheet(_month)

      writeSheetHeader
      writeCompanyDetails
      writeAttendanceHeader
      writeAttendance
      writeHolidays
      writeWeekends

      resizeCols

      using(new FileOutputStream(AttendanceReportFileName))(workbook.write)
    }
  }
}

object ExcelWriter {
  def apply(month: Int, year: Int): ExcelWriter = new ExcelWriter(month, year)
}