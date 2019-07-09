package com.codingkapoor.esslattendancereportgenerator.writer

import java.sql.Date
import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.AttendanceStatus
import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Holiday}
import org.apache.poi.ss.usermodel.{BorderStyle, HorizontalAlignment}
import org.apache.poi.xssf.usermodel.{XSSFColor, XSSFWorkbook}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder

trait AttendanceWriter {
  val month: Int
  val year: Int

  def writeAttendance(implicit workbook: XSSFWorkbook, attendances: Seq[AttendancePerEmployee], holidays: Seq[Holiday]): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString
    val numOfDaysInMonth = yearMonth.lengthOfMonth

    val sheet = workbook.getSheet(_month)

    val createHelper = workbook.getCreationHelper

    val font = workbook.createFont
    font.setFontHeightInPoints(10.toShort)

    val cellStyle = workbook.createCellStyle
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

    val idCellStyle = workbook.createCellStyle
    idCellStyle.setAlignment(HorizontalAlignment.LEFT)

    val dateCellStyle = workbook.createCellStyle
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

    var rowNum = 3
    for (att <- attendances) {
      val employee = att.employee
      val attendance = att.attendance

      val row = sheet.createRow(rowNum)

      val id = row.createCell(0)
      id.setCellValue(employee.empId)
      id.setCellStyle(idCellStyle)
      id.setCellStyle(cellStyle)

      val name = row.createCell(1)
      name.setCellValue(employee.name)
      name.setCellStyle(cellStyle)

      val gender = row.createCell(2)
      gender.setCellValue(employee.gender)
      gender.setCellStyle(cellStyle)

      val doj = row.createCell(3)
      doj.setCellValue(Date.valueOf(employee.doj))
      doj.setCellStyle(dateCellStyle)

      val pfn = row.createCell(4)
      pfn.setCellValue(employee.pfn)
      pfn.setCellStyle(cellStyle)

      var daysIndex = 5
      for (i <- 1 to numOfDaysInMonth) {
        val col = row.createCell(daysIndex)
        if (!attendance(i).equals(AttendanceStatus.Abscond.toString))
          col.setCellValue(attendance(i))
        col.setCellStyle(cellStyle)
        daysIndex += 1
      }

      rowNum += 1
    }
  }
}
