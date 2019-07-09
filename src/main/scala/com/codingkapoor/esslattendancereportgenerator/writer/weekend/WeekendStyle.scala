package com.codingkapoor.esslattendancereportgenerator.writer.weekend

import org.apache.poi.ss.usermodel.{BorderStyle, FillPatternType, HorizontalAlignment, VerticalAlignment}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFColor, XSSFWorkbook}

trait WeekendStyle {

  def getWeekendCellStyle(implicit workbook: XSSFWorkbook): (XSSFCellStyle, XSSFCellStyle) = {
    val font = workbook.createFont
    font.setFontHeightInPoints(10.toShort)

    val satCellStyle: XSSFCellStyle = workbook.createCellStyle
    satCellStyle.setFont(font)
    satCellStyle.setAlignment(HorizontalAlignment.CENTER)
    satCellStyle.setVerticalAlignment(VerticalAlignment.CENTER)
    satCellStyle.setRotation(90.toShort)
    satCellStyle.setShrinkToFit(true)
    satCellStyle.setBorderLeft(BorderStyle.THIN)
    satCellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
    satCellStyle.setBorderTop(BorderStyle.THIN)
    satCellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
    satCellStyle.setBorderRight(BorderStyle.THIN)
    satCellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
    satCellStyle.setBorderBottom(BorderStyle.THIN)
    satCellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
    satCellStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.LIGHT_GRAY))
    satCellStyle.setFillPattern(FillPatternType.DIAMONDS)

    val sunCellStyle: XSSFCellStyle = workbook.createCellStyle
    sunCellStyle.setFont(font)
    sunCellStyle.setAlignment(HorizontalAlignment.CENTER)
    sunCellStyle.setVerticalAlignment(VerticalAlignment.CENTER)
    sunCellStyle.setRotation(90.toShort)
    sunCellStyle.setShrinkToFit(true)
    sunCellStyle.setBorderLeft(BorderStyle.THIN)
    sunCellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
    sunCellStyle.setBorderTop(BorderStyle.THIN)
    sunCellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
    sunCellStyle.setBorderRight(BorderStyle.THIN)
    sunCellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
    sunCellStyle.setBorderBottom(BorderStyle.THIN)
    sunCellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
    sunCellStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.WHITE))
    sunCellStyle.setFillPattern(FillPatternType.DIAMONDS)

    (satCellStyle, sunCellStyle)
  }
}
