package com.codingkapoor.esslattendancereportgenerator.writer.attendanceheader

import com.codingkapoor.esslattendancereportgenerator.writer.AttendanceHeaderDimensions
import com.typesafe.scalalogging.LazyLogging
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait AttendanceHeaderWriter extends AttendanceHeaderStyle with LazyLogging {
  val monthTitle: String

  val attendanceHeaderDimensions: AttendanceHeaderDimensions

  def writeAttendanceHeader(implicit workbook: XSSFWorkbook): Unit = {
    val sheet = workbook.getSheet(monthTitle)

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

    logger.info("writeAttendanceHeader completed.")
  }
}
