package com.codingkapoor.esslattendancereportgenerator.writer.holiday

import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Holiday}
import com.codingkapoor.esslattendancereportgenerator.writer.AttendanceDimensions
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

    val attendanceDimensions = AttendanceDimensions(month, year, attendances.map(l => l.employee))

    val mergedRegionfirstRowIndex = attendanceDimensions.firstRowIndex
    val mergedRegionlastRowIndex = attendanceDimensions.lastRowIndex - 1

    val cellStyle = getHolidayCellStyle

    for (holiday <- holidays) {
      val day = holiday.date.getDayOfMonth
      val dayIndex = attendanceDimensions.firstColumnIndex + day - 1

      val mergedRegionfirstColumnIndex = dayIndex
      val mergedRegionlastColumnIndex = dayIndex

      if (!mergedRegionAlreadyExists(mergedRegionfirstRowIndex, mergedRegionlastRowIndex, mergedRegionfirstColumnIndex, mergedRegionlastColumnIndex)) {

        for (i <- mergedRegionfirstRowIndex to mergedRegionlastRowIndex) {
          val row = sheet.getRow(i)
          for (j <- mergedRegionfirstColumnIndex to mergedRegionlastColumnIndex) {
            val col = row.createCell(j)
            col.setCellStyle(cellStyle)
            if (i == mergedRegionfirstRowIndex && j == mergedRegionfirstColumnIndex) {
              col.setCellValue(holiday.occasion)
            }
          }
        }

        sheet.addMergedRegion(new CellRangeAddress(mergedRegionfirstRowIndex, mergedRegionlastRowIndex, mergedRegionfirstColumnIndex, mergedRegionlastColumnIndex))
      }
    }
  }
}
