package com.codingkapoor.esslattendancereportgenerator.writer.employeeinfoheader

import org.apache.poi.ss.usermodel.{BorderStyle, FillPatternType, HorizontalAlignment, IndexedColors}
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFColor, XSSFWorkbook}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder

trait EmployeeInfoHeaderStyle {

  def getEmployeeInfoHeaderCellStyle(implicit workbook: XSSFWorkbook): XSSFCellStyle = {
    val font = workbook.createFont
    font.setBold(true)
    font.setFontHeightInPoints(10.5.toShort)
    font.setColor(IndexedColors.BLACK.getIndex)

    val cellStyle: XSSFCellStyle = workbook.createCellStyle
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

    cellStyle
  }
}
