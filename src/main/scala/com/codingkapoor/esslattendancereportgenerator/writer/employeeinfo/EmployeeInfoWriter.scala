package com.codingkapoor.esslattendancereportgenerator.writer.employeeinfo

import java.sql.Date

import com.codingkapoor.esslattendancereportgenerator.model.Attendance
import com.codingkapoor.esslattendancereportgenerator.writer.EmployeeInfoDimensions
import com.typesafe.scalalogging.LazyLogging
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait EmployeeInfoWriter extends EmployeeInfoStyle with LazyLogging {
  val monthTitle: String

  val attendances: Seq[Attendance]

  val employeeInfoDimensions: EmployeeInfoDimensions

  def writeEmployeeInfo(implicit workbook: XSSFWorkbook): Unit = {
    val sheet = workbook.getSheet(monthTitle)

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

    logger.info("writeEmployeeInfo completed.")
  }
}
