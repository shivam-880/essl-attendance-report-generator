package com.codingkapoor.esslattendancereportgenerator.writer.holiday

import java.time.LocalDate

import com.codingkapoor.esslattendancereportgenerator.writer.AttendanceDimensions
import com.typesafe.scalalogging.LazyLogging
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait HolidayWriter extends HolidayStyle with LazyLogging {
  val monthTitle: String

  val holidays: Map[LocalDate, String]

  val attendanceDimensions: AttendanceDimensions

  def mergedRegionAlreadyExists(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int)(implicit workbook: XSSFWorkbook): Boolean

  def writeHolidays(implicit workbook: XSSFWorkbook): Unit = {
    val sheet = workbook.getSheet(monthTitle)

    val mergedRegionfirstRowIndex = attendanceDimensions.firstRowIndex
    val mergedRegionlastRowIndex = attendanceDimensions.lastRowIndex - 1

    val cellStyle = getHolidayCellStyle

    val holidaysItr = holidays.iterator
    while (holidaysItr.hasNext) {
      val next = holidaysItr.next
      val date = next._1
      val occasion = next._2

      val day = date.getDayOfMonth
      val dayIndex = attendanceDimensions.firstColumnIndex + day - 1

      val mergedRegionfirstColumnIndex = dayIndex
      val mergedRegionlastColumnIndex = dayIndex

      if (!mergedRegionAlreadyExists(mergedRegionfirstRowIndex, mergedRegionlastRowIndex, mergedRegionfirstColumnIndex, mergedRegionlastColumnIndex)) {

        for (i <- mergedRegionfirstRowIndex to mergedRegionlastRowIndex) {
          val row = if(sheet.getRow(i) != null) sheet.getRow(i) else sheet.createRow(i)

          for (j <- mergedRegionfirstColumnIndex to mergedRegionlastColumnIndex) {
            val col = row.createCell(j)
            col.setCellStyle(cellStyle)
            if (i == mergedRegionfirstRowIndex && j == mergedRegionfirstColumnIndex) {
              col.setCellValue(occasion)
            }
          }
        }

        sheet.addMergedRegion(new CellRangeAddress(mergedRegionfirstRowIndex, mergedRegionlastRowIndex, mergedRegionfirstColumnIndex, mergedRegionlastColumnIndex))
      }
    }

    logger.info("writeHolidays completed.")
  }
}
