package com.codingkapoor.esslattendancereportgenerator.writer.sheetheader

import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.writer.SectionHeaderDimensions
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait SheetHeaderWriter extends SheetHeaderStyle {
  val month: Int
  val year: Int

  def writeSheetHeader(implicit workbook: XSSFWorkbook): Int = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString

    val sheet = workbook.getSheet(_month)

    val sectionHeaderDimensions = SectionHeaderDimensions(month, year)

    val firstRowIndex = sectionHeaderDimensions.firstRowIndex
    val lastRowIndex = sectionHeaderDimensions.lastRowIndex
    val firstColIndex = sectionHeaderDimensions.firstColumnIndex
    val lastColIndex = sectionHeaderDimensions.lastColumnIndex

    val row = sheet.createRow(firstRowIndex)
    row.setHeightInPoints(50)

    val cellStyle = getSheetCellStyle

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
