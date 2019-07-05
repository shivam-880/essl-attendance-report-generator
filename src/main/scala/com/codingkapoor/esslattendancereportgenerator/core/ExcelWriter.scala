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
      val sheet = workbook.createSheet("July")
      val createHelper = workbook.getCreationHelper

      val headerFont = workbook.createFont
      headerFont.setBold(true)
      headerFont.setFontHeightInPoints(10.5.toShort)
      headerFont.setColor(IndexedColors.BLACK.getIndex)

      val headerCellStyle = workbook.createCellStyle
      headerCellStyle.setFont(headerFont)

      val companyDetailsRow = sheet.createRow(0)
      companyDetailsRow.setHeightInPoints(50)
      val companyDetailsCol = companyDetailsRow.createCell(0)
      companyDetailsCol.setCellValue(s"$CompanyName\n$CompanyAddress")
      companyDetailsCol.setCellStyle(headerCellStyle)

      sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3))

      val employeesInfoHeaderRow = sheet.createRow(2)
      for (i <- EmployeesInfoHeader.indices) {
        val employeesInfoHeaderCell = employeesInfoHeaderRow.createCell(i)
        employeesInfoHeaderCell.setCellValue(EmployeesInfoHeader(i))
        employeesInfoHeaderCell.setCellStyle(headerCellStyle)
      }

      val dateCellStyle = workbook.createCellStyle
      dateCellStyle.setDataFormat(createHelper.createDataFormat.getFormat("dd-MM-yyyy"))

      var rowNum = 3
      for (employee <- employees) {
        val employeeInfoRow = sheet.createRow(rowNum)
        employeeInfoRow.createCell(0).setCellValue(employee.name)
        employeeInfoRow.createCell(1).setCellValue(employee.gender)
        val doj = employeeInfoRow.createCell(2)
        doj.setCellValue(employee.doj.toString)
        doj.setCellStyle(dateCellStyle)
        employeeInfoRow.createCell(3).setCellValue(employee.pfn)

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
