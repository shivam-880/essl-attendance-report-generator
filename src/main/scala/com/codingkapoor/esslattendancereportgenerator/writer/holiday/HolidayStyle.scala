package com.codingkapoor.esslattendancereportgenerator.writer.holiday

import org.apache.poi.ss.usermodel.{BorderStyle, FillPatternType, HorizontalAlignment, VerticalAlignment}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFColor, XSSFWorkbook}

trait HolidayStyle {

  def getHolidayCellStyle(implicit workbook: XSSFWorkbook): XSSFCellStyle = {
    val font = workbook.createFont
    font.setFontHeightInPoints(10.toShort)

    val cellStyle: XSSFCellStyle = workbook.createCellStyle
    cellStyle.setFont(font)
    cellStyle.setAlignment(HorizontalAlignment.CENTER)
    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER)
    cellStyle.setRotation(90.toShort)
    cellStyle.setShrinkToFit(true)
    cellStyle.setBorderLeft(BorderStyle.THIN)
    cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
    cellStyle.setBorderTop(BorderStyle.THIN)
    cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
    cellStyle.setBorderRight(BorderStyle.THIN)
    cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
    cellStyle.setBorderBottom(BorderStyle.THIN)
    cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
    cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(238, 204, 251)))
    cellStyle.setFillPattern(FillPatternType.DIAMONDS)

    cellStyle
  }
}
