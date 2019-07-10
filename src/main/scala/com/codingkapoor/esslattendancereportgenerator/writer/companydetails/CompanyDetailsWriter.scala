package com.codingkapoor.esslattendancereportgenerator.writer.companydetails

import com.codingkapoor.esslattendancereportgenerator.`package`.{CompanyAddress, CompanyName}
import com.codingkapoor.esslattendancereportgenerator.writer.CompanyDetailsDimensions
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait CompanyDetailsWriter extends CompanyDetailsStyle {
  val monthTitle: String

  val companyDetailsDimensions: CompanyDetailsDimensions

  def writeCompanyDetails(implicit workbook: XSSFWorkbook): Int = {
    val sheet = workbook.getSheet(monthTitle)

    val firstRowIndex = companyDetailsDimensions.firstColumnIndex
    val lastRowIndex = companyDetailsDimensions.lastRowIndex
    val firstColIndex = companyDetailsDimensions.firstColumnIndex
    val lastColIndex = companyDetailsDimensions.lastColumnIndex

    val row = sheet.getRow(firstRowIndex)
    row.setHeightInPoints(50)

    val cellStyle = getCompanyDetailsCellStyle

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
