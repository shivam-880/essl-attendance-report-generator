package com.codingkapoor.esslattendancereportgenerator.writer.companydetails

import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.`package`.{CompanyAddress, CompanyName}
import org.apache.poi.ss.usermodel.{BorderStyle, FillPatternType, IndexedColors}
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder
import org.apache.poi.xssf.usermodel.{XSSFColor, XSSFWorkbook}

trait CompanyDetailsWriter extends CompanyDetailsStyle {
  val month: Int
  val year: Int

  def writeCompanyDetails(implicit workbook: XSSFWorkbook): Int = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString

    val sheet = workbook.getSheet(_month)

    val cellStyle = getCompanyDetailsCellStyle

    val row = sheet.getRow(0)
    row.setHeightInPoints(50)

    val firstRowIndex = 0
    val lastRowIndex = 0
    val firstColIndex = 0
    val lastColIndex = 4

    for (i <- firstColIndex to lastColIndex) {
      val col = row.createCell(i)
      col.setCellStyle(cellStyle)
      if (i == firstColIndex) {
        col.setCellValue(s"$CompanyName\n$CompanyAddress")
      }
    }

    sheet.addMergedRegion(new CellRangeAddress(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex))
  }
}
