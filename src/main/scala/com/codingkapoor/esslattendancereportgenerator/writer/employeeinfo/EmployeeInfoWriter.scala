package com.codingkapoor.esslattendancereportgenerator.writer.employeeinfo

import java.sql.Date
import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Holiday}
import com.codingkapoor.esslattendancereportgenerator.writer.EmployeeInfoDimensions
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait EmployeeInfoWriter extends EmployeeInfoStyle {
  val month: Int
  val year: Int

  def writeEmployeeInfo(implicit workbook: XSSFWorkbook, attendances: Seq[AttendancePerEmployee], holidays: Seq[Holiday]): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString

    val sheet = workbook.getSheet(_month)

    val employeeInfoDimensions = EmployeeInfoDimensions(attendances.map(l => l.employee))

    val firstRowIndex = employeeInfoDimensions.firstRowIndex
    val firstColumnIndex = employeeInfoDimensions.firstColumnIndex
    val lastColumnIndex = employeeInfoDimensions.lastColumnIndex

    val (cellStyle, idCellStyle, dateCellStyle) = getEmployeeInfoCellStyle

    var rowNum = firstRowIndex
    for (att <- attendances) {
      val employee = att.employee

      val row = sheet.createRow(rowNum)

      val id = row.createCell(firstColumnIndex)
      id.setCellValue(employee.empId)
      id.setCellStyle(idCellStyle)
      id.setCellStyle(cellStyle)

      val name = row.createCell(firstColumnIndex + 1)
      name.setCellValue(employee.name)
      name.setCellStyle(cellStyle)

      val gender = row.createCell(firstColumnIndex + 2)
      gender.setCellValue(employee.gender)
      gender.setCellStyle(cellStyle)

      val doj = row.createCell(firstColumnIndex + 3)
      doj.setCellValue(Date.valueOf(employee.doj))
      doj.setCellStyle(dateCellStyle)

      val pfn = row.createCell(firstColumnIndex + 4)
      pfn.setCellValue(employee.pfn)
      pfn.setCellStyle(cellStyle)

      rowNum += 1
    }

    for (i <- firstColumnIndex to lastColumnIndex)
      sheet.autoSizeColumn(i)
  }
}
