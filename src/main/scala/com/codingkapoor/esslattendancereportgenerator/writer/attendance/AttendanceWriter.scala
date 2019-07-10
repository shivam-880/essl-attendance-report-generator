package com.codingkapoor.esslattendancereportgenerator.writer.attendance

import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.AttendanceStatus
import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Holiday}
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait AttendanceWriter extends AttendanceStyle {
  val month: Int
  val year: Int

  def writeAttendance(implicit workbook: XSSFWorkbook, attendances: Seq[AttendancePerEmployee], holidays: Seq[Holiday]): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString
    val numOfDaysInMonth = yearMonth.lengthOfMonth

    val sheet = workbook.getSheet(_month)

    val cellStyle = getAttendanceCellStyle

    var rowNum = 3
    for (att <- attendances) {
      val attendance = att.attendance

      val row = sheet.getRow(rowNum)

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
