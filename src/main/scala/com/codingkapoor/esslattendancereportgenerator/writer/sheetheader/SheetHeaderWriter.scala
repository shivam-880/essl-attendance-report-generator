package com.codingkapoor.esslattendancereportgenerator.writer.sheetheader

import java.time.YearMonth

import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait SheetHeaderWriter extends SheetHeaderStyle {
  val month: Int
  val year: Int

  def writeSheetHeader(implicit workbook: XSSFWorkbook): Int = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString
    val numOfDaysInMonth = yearMonth.lengthOfMonth

    val sheet = workbook.getSheet(_month)

    val cellStyle = getSheetCellStyle

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
