package com.codingkapoor.esslattendancereportgenerator.writer.attendanceheader

import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.`package`.EmployeesInfoHeader
import org.apache.poi.ss.usermodel.{BorderStyle, FillPatternType, HorizontalAlignment, IndexedColors}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder
import org.apache.poi.xssf.usermodel.{XSSFColor, XSSFWorkbook}

trait AttendanceHeaderWriter extends AttendanceHeaderStyle {
  val month: Int
  val year: Int

  def writeAttendanceHeader(implicit workbook: XSSFWorkbook): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString
    val numOfDaysInMonth = yearMonth.lengthOfMonth

    val sheet = workbook.getSheet(_month)

    val cellStyle = getAttendanceHeaderCellStyle

    val row = sheet.getRow(2)

    var daysIndex = 5
    for (i <- 1 to numOfDaysInMonth) {
      val col = row.createCell(daysIndex)
      col.setCellValue(i)
      col.setCellStyle(cellStyle)
      daysIndex += 1
    }
  }
}
