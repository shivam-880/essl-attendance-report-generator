package com.codingkapoor.esslattendancereportgenerator.writer.attendance

import java.sql.Date
import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.AttendanceStatus
import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Holiday}
import org.apache.poi.ss.usermodel.{BorderStyle, HorizontalAlignment}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder
import org.apache.poi.xssf.usermodel.{XSSFColor, XSSFWorkbook}

trait AttendanceWriter extends AttendanceStyle {
  val month: Int
  val year: Int

  def writeAttendance(implicit workbook: XSSFWorkbook, attendances: Seq[AttendancePerEmployee], holidays: Seq[Holiday]): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString
    val numOfDaysInMonth = yearMonth.lengthOfMonth

    val sheet = workbook.getSheet(_month)

    val (cellStyle, idCellStyle, dateCellStyle) = getAttendanceCellStyle

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
