package com.codingkapoor.esslattendancereportgenerator.writer.attendance

import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.AttendanceStatus
import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Holiday}
import com.codingkapoor.esslattendancereportgenerator.writer.AttendanceDimensions
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait AttendanceWriter extends AttendanceStyle {
  val month: Int
  val year: Int

  def writeAttendance(implicit workbook: XSSFWorkbook, attendances: Seq[AttendancePerEmployee], holidays: Seq[Holiday]): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString

    val sheet = workbook.getSheet(_month)

    val attendanceDimensions = AttendanceDimensions(month, year, attendances.map(l => l.employee))

    val firstRowIndex = attendanceDimensions.firstRowIndex
    val firstColumnIndex = attendanceDimensions.firstColumnIndex
    val lastColumnIndex = attendanceDimensions.lastColumnIndex

    val cellStyle = getAttendanceCellStyle

    var rowNum = firstRowIndex
    for (att <- attendances) {
      val attendance = att.attendance

      val row = sheet.getRow(rowNum)

      var day = 1
      for (i <- firstColumnIndex to lastColumnIndex) {
        val col = row.createCell(i)
        if (!attendance(day).equals(AttendanceStatus.Abscond.toString))
          col.setCellValue(attendance(day))
        col.setCellStyle(cellStyle)
        day += 1
      }

      rowNum += 1
    }

    for (i <- firstColumnIndex to lastColumnIndex)
      sheet.setColumnWidth(i, 1200)
  }
}
