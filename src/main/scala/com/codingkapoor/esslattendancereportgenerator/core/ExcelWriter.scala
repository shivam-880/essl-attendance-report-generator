package com.codingkapoor.esslattendancereportgenerator.core

import java.io.FileOutputStream
import java.sql.Date
import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.model.Employee
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel.{HorizontalAlignment, IndexedColors}
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

      val row = sheet.createRow(0)
      row.setHeightInPoints(50)

      val col = row.createCell(0)
      col.setCellValue(s"$CompanyName\n$CompanyAddress")
      col.setCellStyle(cellStyle)

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

      val row = sheet.createRow(2)

      for (i <- EmployeesInfoHeader.indices) {
        val col = row.createCell(i)
        col.setCellValue(EmployeesInfoHeader(i))
        col.setCellStyle(cellStyle)
      }

      var daysIndex = 5
      for (i <- 1 to numOfDays) {
        val col = row.createCell(daysIndex)
        col.setCellValue(i)
        col.setCellStyle(cellStyle)
        daysIndex += 1
      }
    }

    def writeAttendance(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)

      val createHelper = workbook.getCreationHelper

      val dateCellStyle = workbook.createCellStyle
      dateCellStyle.setDataFormat(createHelper.createDataFormat.getFormat("dd-mmm-yyyy"))
      dateCellStyle.setAlignment(HorizontalAlignment.LEFT)

      val idCellStyle = workbook.createCellStyle
      idCellStyle.setAlignment(HorizontalAlignment.LEFT)

      var rowNum = 3
      for (employee <- employees) {
        val row = sheet.createRow(rowNum)

        val id = row.createCell(0)
        id.setCellValue(employee.empId)
        id.setCellStyle(idCellStyle)

        val name = row.createCell(1)
        name.setCellValue(employee.name)

        val gender = row.createCell(2)
        gender.setCellValue(employee.gender)

        val doj = row.createCell(3)
        doj.setCellValue(Date.valueOf(employee.doj))
        doj.setCellStyle(dateCellStyle)

        val pfn = row.createCell(4)
        pfn.setCellValue(employee.pfn)

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

      using(new FileOutputStream(AttendanceReportFileName))(workbook.write)
    }
  }
}
