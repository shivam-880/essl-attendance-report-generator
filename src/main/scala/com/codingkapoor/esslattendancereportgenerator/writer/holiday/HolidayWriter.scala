package com.codingkapoor.esslattendancereportgenerator.writer.holiday

import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Holiday}
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait HolidayWriter extends HolidayStyle {
  val month: Int
  val year: Int

  def mergedRegionAlreadyExists(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int)(implicit workbook: XSSFWorkbook): Boolean

  def writeHolidays(implicit workbook: XSSFWorkbook, attendances: Seq[AttendancePerEmployee], holidays: Seq[Holiday]): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString

    val sheet = workbook.getSheet(_month)

    val cellStyle = getHolidayCellStyle

    for (holiday <- holidays) {
      val day = holiday.date.getDayOfMonth
      val dayIndex = 5 + day - 1

      val firstRowIndex = 3
      val lastRowIndex = firstRowIndex + attendances.size - 1
      val firstColIndex = dayIndex
      val lastColIndex = dayIndex

      if (!mergedRegionAlreadyExists(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex)) {

        for (i <- firstRowIndex to lastRowIndex) {
          val row = sheet.getRow(i)
          for (j <- firstColIndex to lastColIndex) {
            val col = row.createCell(j)
            col.setCellStyle(cellStyle)
            if (i == firstRowIndex && j == firstColIndex) {
              col.setCellValue(holiday.occasion)
            }
          }
        }

        sheet.addMergedRegion(new CellRangeAddress(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex))
      }
    }
  }
}
