package com.codingkapoor.esslattendancereportgenerator.writer.weekend

import java.time.{DayOfWeek, LocalDate, YearMonth}

import com.codingkapoor.esslattendancereportgenerator.writer.AttendanceDimensions
import com.typesafe.scalalogging.LazyLogging
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait WeekendWriter extends WeekendStyle with LazyLogging {
  val month: Int
  val year: Int
  val yearMonth: YearMonth

  val monthTitle: String

  val attendanceDimensions: AttendanceDimensions

  def mergedRegionAlreadyExists(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int)(implicit workbook: XSSFWorkbook): Boolean

  def writeWeekends(implicit workbook: XSSFWorkbook): Unit = {
    val sheet = workbook.getSheet(monthTitle)

    val mergedRegionFirstRowIndex = attendanceDimensions.firstRowIndex
    val mergedRegionLastRowIndex = attendanceDimensions.lastRowIndex - 1

    val (satCellStyle, sunCellStyle) = getWeekendCellStyle

    for (dayOfMonth <- 1 to yearMonth.lengthOfMonth()) {
      val dayOfWeek = LocalDate.of(year, month, dayOfMonth).getDayOfWeek
      if (dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY) {
        val dayIndex = attendanceDimensions.firstColumnIndex + dayOfMonth - 1

        val mergedRegionFirstColumnIndex = dayIndex
        val mergedRegionLastColumnIndex = dayIndex

        if (!mergedRegionAlreadyExists(mergedRegionFirstRowIndex, mergedRegionLastRowIndex, mergedRegionFirstColumnIndex, mergedRegionLastColumnIndex)) {
          for (i <- mergedRegionFirstRowIndex to mergedRegionLastRowIndex) {
            val row = sheet.getRow(i)
            for (j <- mergedRegionFirstColumnIndex to mergedRegionLastColumnIndex) {
              val col = row.createCell(j)

              if (dayOfWeek == DayOfWeek.SATURDAY) col.setCellStyle(satCellStyle)
              else col.setCellStyle(sunCellStyle)

              if (i == mergedRegionFirstRowIndex && j == mergedRegionFirstColumnIndex) {
                val dayOfWeekStr = dayOfWeek.toString
                val camelCasedDayOfWeek = dayOfWeekStr.take(1) + dayOfWeekStr.drop(1).toLowerCase
                col.setCellValue(camelCasedDayOfWeek)
              }
            }
          }

          sheet.addMergedRegion(new CellRangeAddress(mergedRegionFirstRowIndex, mergedRegionLastRowIndex, mergedRegionFirstColumnIndex, mergedRegionLastColumnIndex))
        }
      }
    }

    logger.info("writeWeekends completed.")
  }
}
