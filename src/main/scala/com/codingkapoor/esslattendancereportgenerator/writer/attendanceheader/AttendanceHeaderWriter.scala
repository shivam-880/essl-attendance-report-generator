package com.codingkapoor.esslattendancereportgenerator.writer.attendanceheader

import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.writer.AttendanceHeaderDimensions
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait AttendanceHeaderWriter extends AttendanceHeaderStyle {
  val month: Int
  val year: Int

  def writeAttendanceHeader(implicit workbook: XSSFWorkbook): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString

    val sheet = workbook.getSheet(_month)

    val attendanceHeaderDimensions = AttendanceHeaderDimensions(month, year)

    val firstRowIndex = attendanceHeaderDimensions.firstRowIndex
    val firstColumnIndex = attendanceHeaderDimensions.firstColumnIndex
    val lastColumnIndex = attendanceHeaderDimensions.lastColumnIndex

    val row = sheet.getRow(firstRowIndex)

    val cellStyle = getAttendanceHeaderCellStyle

    var day = 1
    for (i <- firstColumnIndex to lastColumnIndex) {
      val col = row.createCell(i)
      col.setCellValue(day)
      col.setCellStyle(cellStyle)
      day += 1
    }
  }
}
