package com.codingkapoor.esslattendancereportgenerator.writer

import java.time.YearMonth

import org.apache.poi.ss.usermodel._
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.{XSSFColor, XSSFWorkbook}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder

trait SheetHeaderWriter {
  val month: Int
  val year: Int

  def writeSheetHeader(implicit workbook: XSSFWorkbook): Int = {
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
    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER)
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

    val row = sheet.createRow(0)
    row.setHeightInPoints(50)

    val firstRowIndex = 0
    val lastRowIndex = 0
    val firstColIndex = 5
    val lastColIndex = 5 + numOfDaysInMonth - 1

    for (i <- firstColIndex to lastColIndex) {
      val col = row.createCell(i)
      col.setCellStyle(cellStyle)
      if (i == firstColIndex) {
        col.setCellValue(s"${_month}, $year")
      }
    }

    sheet.addMergedRegion(new CellRangeAddress(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex))
  }
}
