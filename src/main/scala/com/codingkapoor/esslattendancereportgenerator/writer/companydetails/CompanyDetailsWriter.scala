package com.codingkapoor.esslattendancereportgenerator.writer.companydetails

import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.`package`.{CompanyAddress, CompanyName}
import com.codingkapoor.esslattendancereportgenerator.writer.CompanyDetailsDimensions
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait CompanyDetailsWriter extends CompanyDetailsStyle {
  val month: Int
  val year: Int

  def writeCompanyDetails(implicit workbook: XSSFWorkbook): Int = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString

    val sheet = workbook.getSheet(_month)

    val companyDetailsDimensions = CompanyDetailsDimensions()

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
