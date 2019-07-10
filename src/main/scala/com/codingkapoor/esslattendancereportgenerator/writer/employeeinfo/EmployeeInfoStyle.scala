package com.codingkapoor.esslattendancereportgenerator.writer.employeeinfo

import org.apache.poi.ss.usermodel.{BorderStyle, HorizontalAlignment}
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFColor, XSSFWorkbook}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder

trait EmployeeInfoStyle {

  def getEmployeeInfoCellStyle(implicit workbook: XSSFWorkbook): (XSSFCellStyle, XSSFCellStyle, XSSFCellStyle) = {
    val font = workbook.createFont
    font.setFontHeightInPoints(10.toShort)

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

    val idCellStyle: XSSFCellStyle = workbook.createCellStyle
    idCellStyle.setAlignment(HorizontalAlignment.LEFT)

    val createHelper = workbook.getCreationHelper
    val dateCellStyle: XSSFCellStyle = workbook.createCellStyle
    dateCellStyle.setDataFormat(createHelper.createDataFormat.getFormat("dd-mmm-yyyy"))
    dateCellStyle.setAlignment(HorizontalAlignment.CENTER)
    dateCellStyle.setBorderLeft(BorderStyle.THIN)
    dateCellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
    dateCellStyle.setBorderTop(BorderStyle.THIN)
    dateCellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
    dateCellStyle.setBorderRight(BorderStyle.THIN)
    dateCellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
    dateCellStyle.setBorderBottom(BorderStyle.THIN)
    dateCellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))

    (cellStyle, idCellStyle, dateCellStyle)
  }
}
