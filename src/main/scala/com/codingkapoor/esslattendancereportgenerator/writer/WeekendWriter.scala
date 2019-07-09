package com.codingkapoor.esslattendancereportgenerator.writer

import java.time.{DayOfWeek, LocalDate, YearMonth}

import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Holiday}
import org.apache.poi.ss.usermodel.{BorderStyle, FillPatternType, HorizontalAlignment, VerticalAlignment}
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.{XSSFColor, XSSFWorkbook}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder

trait WeekendWriter {
  val month: Int
  val year: Int

  def mergedRegionAlreadyExists(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int)(implicit workbook: XSSFWorkbook): Boolean

  def writeWeekends(implicit workbook: XSSFWorkbook, attendances: Seq[AttendancePerEmployee], holidays: Seq[Holiday]): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val _month = yearMonth.getMonth.toString

    val sheet = workbook.getSheet(_month)

    val font = workbook.createFont
    font.setFontHeightInPoints(10.toShort)

    val satCellStyle = workbook.createCellStyle
    satCellStyle.setFont(font)
    satCellStyle.setAlignment(HorizontalAlignment.CENTER)
    satCellStyle.setVerticalAlignment(VerticalAlignment.CENTER)
    satCellStyle.setRotation(90.toShort)
    satCellStyle.setShrinkToFit(true)
    satCellStyle.setBorderLeft(BorderStyle.THIN)
    satCellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
    satCellStyle.setBorderTop(BorderStyle.THIN)
    satCellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
    satCellStyle.setBorderRight(BorderStyle.THIN)
    satCellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
    satCellStyle.setBorderBottom(BorderStyle.THIN)
    satCellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
    satCellStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.LIGHT_GRAY))
    satCellStyle.setFillPattern(FillPatternType.DIAMONDS)

    val sunCellStyle = workbook.createCellStyle
    sunCellStyle.setFont(font)
    sunCellStyle.setAlignment(HorizontalAlignment.CENTER)
    sunCellStyle.setVerticalAlignment(VerticalAlignment.CENTER)
    sunCellStyle.setRotation(90.toShort)
    sunCellStyle.setShrinkToFit(true)
    sunCellStyle.setBorderLeft(BorderStyle.THIN)
    sunCellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
    sunCellStyle.setBorderTop(BorderStyle.THIN)
    sunCellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
    sunCellStyle.setBorderRight(BorderStyle.THIN)
    sunCellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
    sunCellStyle.setBorderBottom(BorderStyle.THIN)
    sunCellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
    sunCellStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.WHITE))
    sunCellStyle.setFillPattern(FillPatternType.DIAMONDS)

    for (dayOfMonth <- 1 to yearMonth.lengthOfMonth()) {
      val dayOfWeek = LocalDate.of(year, month, dayOfMonth).getDayOfWeek
      if (dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY) {
        val dayIndex = 5 + dayOfMonth - 1

        val firstRowIndex = 3
        val lastRowIndex = firstRowIndex + attendances.size - 1
        val firstColIndex = dayIndex
        val lastColIndex = dayIndex

        if (!mergedRegionAlreadyExists(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex)) {
          for (i <- firstRowIndex to lastRowIndex) {
            val row = sheet.getRow(i)
            for (j <- firstColIndex to lastColIndex) {
              val col = row.createCell(j)

              if (dayOfWeek == DayOfWeek.SATURDAY) col.setCellStyle(satCellStyle)
              else col.setCellStyle(sunCellStyle)

              if (i == firstRowIndex && j == firstColIndex) {
                val dayOfWeekStr = dayOfWeek.toString
                val camelCasedDayOfWeek = dayOfWeekStr.take(1) + dayOfWeekStr.drop(1).toLowerCase
                col.setCellValue(camelCasedDayOfWeek)
              }
            }
          }

          sheet.addMergedRegion(new CellRangeAddress(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex))
        }
      }
    }
  }
}
