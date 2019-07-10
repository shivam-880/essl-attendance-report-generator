package com.codingkapoor.esslattendancereportgenerator.writer.employeeinfo

import java.sql.Date
import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.`package`.EmployeesInfoHeader
import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Holiday}
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait EmployeeInfoWriter extends EmployeeInfoStyle {
  val month: Int
  val year: Int

  def writeEmployeeInfo(implicit workbook: XSSFWorkbook, attendances: Seq[AttendancePerEmployee], holidays: Seq[Holiday]): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString

    val sheet = workbook.getSheet(_month)

    val (cellStyle, idCellStyle, dateCellStyle) = getEmployeeInfoCellStyle

    var rowNum = 3
    for (att <- attendances) {
      val employee = att.employee

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

      rowNum += 1
    }

    for (i <- EmployeesInfoHeader.indices)
      sheet.autoSizeColumn(i)
  }
}
