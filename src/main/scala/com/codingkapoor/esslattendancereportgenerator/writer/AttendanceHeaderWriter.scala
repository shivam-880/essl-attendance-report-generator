package com.codingkapoor.esslattendancereportgenerator.writer

import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.`package`.EmployeesInfoHeader
import org.apache.poi.ss.usermodel.{BorderStyle, FillPatternType, HorizontalAlignment, IndexedColors}
import org.apache.poi.xssf.usermodel.{XSSFColor, XSSFWorkbook}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder

trait AttendanceHeaderWriter {
  val month: Int
  val year: Int

  def writeAttendanceHeader(implicit workbook: XSSFWorkbook): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString
    val numOfDaysInMonth = yearMonth.lengthOfMonth

    val sheet = workbook.getSheet(_month)

    val font = workbook.createFont
    font.setBold(true)
    font.setFontHeightInPoints(10.5.toShort)
    font.setColor(IndexedColors.BLACK.getIndex)

    val cellStyle = workbook.createCellStyle
    cellStyle.setFont(font)
    cellStyle.setAlignment(HorizontalAlignment.CENTER)
    cellStyle.setBorderLeft(BorderStyle.THIN)
    cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
    cellStyle.setBorderTop(BorderStyle.THIN)
    cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
    cellStyle.setBorderRight(BorderStyle.THIN)
    cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
    cellStyle.setBorderBottom(BorderStyle.THIN)
    cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
    cellStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.LIGHT_GRAY))
    cellStyle.setFillPattern(FillPatternType.DIAMONDS)

    val row = sheet.createRow(2)

    for (i <- EmployeesInfoHeader.indices) {
      val col = row.createCell(i)
      col.setCellValue(EmployeesInfoHeader(i))
      col.setCellStyle(cellStyle)
    }

    var daysIndex = 5
    for (i <- 1 to numOfDaysInMonth) {
      val col = row.createCell(daysIndex)
      col.setCellValue(i)
      col.setCellStyle(cellStyle)
      daysIndex += 1
    }
  }
}
