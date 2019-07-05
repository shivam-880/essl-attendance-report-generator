package com.codingkapoor.esslattendancereportgenerator.core

import java.io.FileOutputStream
import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.model.Employee
import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.util.CellRangeAddress
import com.codingkapoor.esslattendancereportgenerator._

object ExcelWriter {

  def write(month: Int, year: Int, employees: Seq[Employee]): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val monthStr = yearMonth.getMonth.toString
    val numOfDays = yearMonth.lengthOfMonth

    def writeCompanyDetails(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)

      val font = workbook.createFont
      font.setBold(true)
      font.setFontHeightInPoints(10.5.toShort)
      font.setColor(IndexedColors.BLACK.getIndex)

      val cellStyle = workbook.createCellStyle
      cellStyle.setFont(font)

      val companyDetailsRow = sheet.createRow(0)
      companyDetailsRow.setHeightInPoints(50)
      val companyDetailsCol = companyDetailsRow.createCell(0)
      companyDetailsCol.setCellValue(s"$CompanyName\n$CompanyAddress")
      companyDetailsCol.setCellStyle(cellStyle)

      sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3))
    }

    def writeAttendanceHeader(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)

      val font = workbook.createFont
      font.setBold(true)
      font.setFontHeightInPoints(10.5.toShort)
      font.setColor(IndexedColors.BLACK.getIndex)

      val cellStyle = workbook.createCellStyle
      cellStyle.setFont(font)

      val employeesInfoHeaderRow = sheet.createRow(2)
      for (i <- EmployeesInfoHeader.indices) {
        val employeesInfoHeaderCell = employeesInfoHeaderRow.createCell(i)
        employeesInfoHeaderCell.setCellValue(EmployeesInfoHeader(i))
        employeesInfoHeaderCell.setCellStyle(cellStyle)
      }
    }

    def writeAttendance(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)

      val createHelper = workbook.getCreationHelper
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
    }

    using(new XSSFWorkbook) { implicit workbook =>
      val sheet = workbook.createSheet(monthStr)

      writeCompanyDetails
      writeAttendanceHeader
      writeAttendance

      for (i <- EmployeesInfoHeader.indices)
        sheet.autoSizeColumn(i)

      using(new FileOutputStream(AttendanceReportFileName)) { os =>
        workbook.write(os)
      }
    }
  }
}
