package com.codingkapoor.esslattendancereportgenerator.core

import java.io.FileOutputStream

import com.codingkapoor.esslattendancereportgenerator.model.Employee
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.util.CellRangeAddress
import com.codingkapoor.esslattendancereportgenerator._

object ExcelWriter {

  def write(employees: Seq[Employee]): Unit = {

    using(new XSSFWorkbook) { workbook =>

      val createHelper = workbook.getCreationHelper

      val sheet = workbook.createSheet("July")

      val headerFont = workbook.createFont
      headerFont.setBold(true)
      headerFont.setFontHeightInPoints(11.toShort)
      headerFont.setColor(IndexedColors.BLACK.getIndex)

      val headerCellStyle = workbook.createCellStyle
      headerCellStyle.setFont(headerFont)

      val headerRow1 = sheet.createRow(0)
      headerRow1.setHeightInPoints(50)
      val cell1 = headerRow1.createCell(0)
      cell1.setCellValue(s"$CompanyName\n$CompanyAddress")
      cell1.setCellStyle(headerCellStyle)

      sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4))

      val headerRow = sheet.createRow(2)

      for (i <- EmployeesInfoHeader.indices) {
        val cell = headerRow.createCell(i)
        cell.setCellValue(EmployeesInfoHeader(i))
        cell.setCellStyle(headerCellStyle)
      }

      val dateCellStyle = workbook.createCellStyle
      dateCellStyle.setDataFormat(createHelper.createDataFormat.getFormat("dd-MM-yyyy"))

      var rowNum = 3
      for (employee <- employees) {
        val row = sheet.createRow(rowNum)
        row.createCell(0).setCellValue(employee.name)
        row.createCell(1).setCellValue(employee.gender)
        val doj = row.createCell(2)
        doj.setCellValue(employee.doj.toString)
        doj.setCellStyle(dateCellStyle)
        row.createCell(3).setCellValue(employee.pfn)

        rowNum += 1
      }

      for (i <- EmployeesInfoHeader.indices)
        sheet.autoSizeColumn(i)

      using(new FileOutputStream(AttendanceReportFileName)) { os =>
        workbook.write(os)
      }
    }
  }
}
