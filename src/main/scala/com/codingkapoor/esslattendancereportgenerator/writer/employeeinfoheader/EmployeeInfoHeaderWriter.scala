package com.codingkapoor.esslattendancereportgenerator.writer.employeeinfoheader

import com.codingkapoor.esslattendancereportgenerator.`package`.EmployeesInfoHeader
import com.codingkapoor.esslattendancereportgenerator.writer.EmployeeInfoHeaderDimensions
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait EmployeeInfoHeaderWriter extends EmployeeInfoHeaderStyle {
  val monthTitle: String

  val employeeInfoHeaderDimensions: EmployeeInfoHeaderDimensions

  def writeEmployeeInfoHeader(implicit workbook: XSSFWorkbook): Unit = {
    val sheet = workbook.getSheet(monthTitle)

    val firstRowIndex = employeeInfoHeaderDimensions.firstRowIndex
    val firstColumnIndex = employeeInfoHeaderDimensions.firstColumnIndex
    val lastColumnIndex = employeeInfoHeaderDimensions.lastColumnIndex

    val row = sheet.createRow(firstRowIndex)

    val cellStyle = getEmployeeInfoHeaderCellStyle

    for (i <- firstColumnIndex to lastColumnIndex) {
      val col = row.createCell(i)
      col.setCellValue(EmployeesInfoHeader(i))
      col.setCellStyle(cellStyle)
    }
  }
}
