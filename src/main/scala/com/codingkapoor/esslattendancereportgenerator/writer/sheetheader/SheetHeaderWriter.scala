package com.codingkapoor.esslattendancereportgenerator.writer.sheetheader

import com.codingkapoor.esslattendancereportgenerator.writer.SectionHeaderDimensions
import com.typesafe.scalalogging.LazyLogging
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait SheetHeaderWriter extends SheetHeaderStyle with LazyLogging {
  val year: Int
  val monthTitle: String

  val sectionHeaderDimensions: SectionHeaderDimensions

  def writeSheetHeader(implicit workbook: XSSFWorkbook): Unit = {
    val sheet = workbook.getSheet(monthTitle)

    val firstRowIndex = sectionHeaderDimensions.firstRowIndex
    val lastRowIndex = sectionHeaderDimensions.lastRowIndex
    val firstColIndex = sectionHeaderDimensions.firstColumnIndex
    val lastColIndex = sectionHeaderDimensions.lastColumnIndex

    val row = if(sheet.getRow(firstRowIndex) != null) sheet.getRow(firstRowIndex) else sheet.createRow(firstRowIndex)
    row.setHeightInPoints(50)

    val cellStyle = getSheetCellStyle

    for (i <- firstColIndex to lastColIndex) {
      val col = row.createCell(i)
      col.setCellStyle(cellStyle)
      if (i == firstColIndex) {
        col.setCellValue(s"$monthTitle, $year")
      }
    }

    sheet.addMergedRegion(new CellRangeAddress(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex))

    logger.info("writeSheetHeader completed.")
  }
}
