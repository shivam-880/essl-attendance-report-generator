package com.codingkapoor.esslattendancereportgenerator.core

import java.io.FileOutputStream
import java.sql.Date
import java.time.{DayOfWeek, LocalDate, YearMonth}

import com.codingkapoor.esslattendancereportgenerator.model.{AttendancePerEmployee, Employee, Holiday}
import org.apache.poi.xssf.usermodel.{XSSFColor, XSSFWorkbook}
import org.apache.poi.ss.usermodel._
import org.apache.poi.ss.util.CellRangeAddress
import com.codingkapoor.esslattendancereportgenerator._
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder

object ExcelWriter {

  def write(attendances: Seq[AttendancePerEmployee], holidays: Seq[Holiday])(month: Int, year: Int): Unit = {
    val yearMonth = YearMonth.of(year, month)
    val monthStr = yearMonth.getMonth.toString
    val numOfDays = yearMonth.lengthOfMonth

    def writeSheetHeader(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)

      val font = workbook.createFont
      font.setBold(true)
      font.setFontHeightInPoints(10.5.toShort)
      font.setColor(IndexedColors.BLACK.getIndex)

      val cellStyle = workbook.createCellStyle
      cellStyle.setFont(font)
      cellStyle.setAlignment(HorizontalAlignment.CENTER)
      cellStyle.setVerticalAlignment(VerticalAlignment.CENTER)
      cellStyle.setBorderLeft(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderTop(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderRight(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderBottom(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.LIGHT_GRAY))
      cellStyle.setFillPattern(FillPatternType.DIAMONDS)

      val row = sheet.createRow(0)
      row.setHeightInPoints(50)

      val firstRowIndex = 0
      val lastRowIndex = 0
      val firstColIndex = 5
      val lastColIndex = 5 + numOfDays - 1

      for (i <- firstColIndex to lastColIndex) {
        val col = row.createCell(i)
        col.setCellStyle(cellStyle)
        if (i == firstColIndex) {
          col.setCellValue(s"$monthStr, $year")
        }
      }

      sheet.addMergedRegion(new CellRangeAddress(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex))
    }

    def writeCompanyDetails(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)

      val font = workbook.createFont
      font.setBold(true)
      font.setFontHeightInPoints(10.5.toShort)
      font.setColor(IndexedColors.BLACK.getIndex)

      val cellStyle = workbook.createCellStyle
      cellStyle.setFont(font)
      cellStyle.setBorderLeft(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderTop(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderRight(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderBottom(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.LIGHT_GRAY))
      cellStyle.setFillPattern(FillPatternType.DIAMONDS)

      val row = sheet.getRow(0)
      row.setHeightInPoints(50)

      val firstRowIndex = 0
      val lastRowIndex = 0
      val firstColIndex = 0
      val lastColIndex = 4

      for (i <- firstColIndex to lastColIndex) {
        val col = row.createCell(i)
        col.setCellStyle(cellStyle)
        if (i == firstColIndex) {
          col.setCellValue(s"$CompanyName\n$CompanyAddress")
        }
      }

      sheet.addMergedRegion(new CellRangeAddress(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex))
    }

    def writeAttendanceHeader(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)

      val font = workbook.createFont
      font.setBold(true)
      font.setFontHeightInPoints(10.5.toShort)
      font.setColor(IndexedColors.BLACK.getIndex)

      val cellStyle = workbook.createCellStyle
      cellStyle.setFont(font)
      cellStyle.setAlignment(HorizontalAlignment.CENTER)
      cellStyle.setBorderLeft(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderTop(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderRight(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderBottom(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.LIGHT_GRAY))
      cellStyle.setFillPattern(FillPatternType.DIAMONDS)

      val row = sheet.createRow(2)

      for (i <- EmployeesInfoHeader.indices) {
        val col = row.createCell(i)
        col.setCellValue(EmployeesInfoHeader(i))
        col.setCellStyle(cellStyle)
      }

      var daysIndex = 5
      for (i <- 1 to numOfDays) {
        val col = row.createCell(daysIndex)
        col.setCellValue(i)
        col.setCellStyle(cellStyle)
        daysIndex += 1
      }
    }

    def writeAttendance(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)

      val createHelper = workbook.getCreationHelper

      val font = workbook.createFont
      font.setFontHeightInPoints(10.toShort)

      val cellStyle = workbook.createCellStyle
      cellStyle.setFont(font)
      cellStyle.setAlignment(HorizontalAlignment.CENTER)
      cellStyle.setBorderLeft(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderTop(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderRight(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderBottom(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))

      val idCellStyle = workbook.createCellStyle
      idCellStyle.setAlignment(HorizontalAlignment.LEFT)

      val dateCellStyle = workbook.createCellStyle
      dateCellStyle.setDataFormat(createHelper.createDataFormat.getFormat("dd-mmm-yyyy"))
      dateCellStyle.setAlignment(HorizontalAlignment.CENTER)
      dateCellStyle.setBorderLeft(BorderStyle.THIN)
      dateCellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
      dateCellStyle.setBorderTop(BorderStyle.THIN)
      dateCellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
      dateCellStyle.setBorderRight(BorderStyle.THIN)
      dateCellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
      dateCellStyle.setBorderBottom(BorderStyle.THIN)
      dateCellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))

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
        for (i <- 1 to numOfDays) {
          val col = row.createCell(daysIndex)
          if (!attendance(i).equals(AttendanceStatus.Abscond.toString))
            col.setCellValue(attendance(i))
          col.setCellStyle(cellStyle)
          daysIndex += 1
        }

        rowNum += 1
      }
    }

    def mergedRegionAlreadyExists(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int)(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)

      var existingMergedRegionsFound = List.empty[Boolean]
      val mergedRegions = sheet.getMergedRegions.iterator()
      while (mergedRegions.hasNext) {
        val cellRangeAddr = mergedRegions.next()
        if (cellRangeAddr.getFirstRow == firstRowIndex && cellRangeAddr.getLastRow == lastRowIndex &&
          cellRangeAddr.getFirstColumn == firstColumnIndex && cellRangeAddr.getLastColumn == lastColumnIndex)
          existingMergedRegionsFound = true :: existingMergedRegionsFound
      }

      existingMergedRegionsFound.contains(true)
    }

    def writeHolidays(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)
      val row = sheet.getRow(3)

      val font = workbook.createFont
      font.setFontHeightInPoints(10.toShort)

      val cellStyle = workbook.createCellStyle
      cellStyle.setFont(font)
      cellStyle.setAlignment(HorizontalAlignment.CENTER)
      cellStyle.setVerticalAlignment(VerticalAlignment.CENTER)
      cellStyle.setRotation(90.toShort)
      cellStyle.setShrinkToFit(true)
      cellStyle.setBorderLeft(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderTop(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderRight(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setBorderBottom(BorderStyle.THIN)
      cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
      cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(238, 204, 251)))
      cellStyle.setFillPattern(FillPatternType.DIAMONDS)

      for (holiday <- holidays) {
        val day = holiday.date.getDayOfMonth
        val dayIndex = 5 + day - 1

        if (!mergedRegionAlreadyExists(3, 20, dayIndex, dayIndex)) {
          val col = row.createCell(dayIndex)
          col.setCellValue(holiday.occasion)
          col.setCellStyle(cellStyle)

          sheet.addMergedRegion(new CellRangeAddress(3, 20, dayIndex, dayIndex))
        }
      }
    }

    def writeWeekends(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)
      val row = sheet.getRow(3)

      val font = workbook.createFont
      font.setFontHeightInPoints(10.toShort)

      for (dayOfMonth <- 1 to yearMonth.lengthOfMonth()) {
        val dayOfWeek = LocalDate.of(year, month, dayOfMonth).getDayOfWeek
        if (dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY) {
          val dayIndex = 5 + dayOfMonth - 1

          if (!mergedRegionAlreadyExists(3, 20, dayIndex, dayIndex)) {
            if (dayOfWeek == DayOfWeek.SATURDAY) {
              val cellStyle = workbook.createCellStyle
              cellStyle.setFont(font)
              cellStyle.setAlignment(HorizontalAlignment.CENTER)
              cellStyle.setVerticalAlignment(VerticalAlignment.CENTER)
              cellStyle.setRotation(90.toShort)
              cellStyle.setShrinkToFit(true)
              cellStyle.setBorderLeft(BorderStyle.THIN)
              cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
              cellStyle.setBorderTop(BorderStyle.THIN)
              cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
              cellStyle.setBorderRight(BorderStyle.THIN)
              cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
              cellStyle.setBorderBottom(BorderStyle.THIN)
              cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
              cellStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.LIGHT_GRAY))
              cellStyle.setFillPattern(FillPatternType.DIAMONDS)

              val col = row.createCell(dayIndex)
              val dayOfWeekStr = dayOfWeek.toString
              col.setCellValue(dayOfWeekStr.take(1) + dayOfWeekStr.drop(1).toLowerCase)
              col.setCellStyle(cellStyle)
            } else {
              val cellStyle = workbook.createCellStyle
              cellStyle.setFont(font)
              cellStyle.setAlignment(HorizontalAlignment.CENTER)
              cellStyle.setVerticalAlignment(VerticalAlignment.CENTER)
              cellStyle.setRotation(90.toShort)
              cellStyle.setShrinkToFit(true)
              cellStyle.setBorderLeft(BorderStyle.THIN)
              cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(java.awt.Color.GRAY))
              cellStyle.setBorderTop(BorderStyle.THIN)
              cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(java.awt.Color.GRAY))
              cellStyle.setBorderRight(BorderStyle.THIN)
              cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(java.awt.Color.GRAY))
              cellStyle.setBorderBottom(BorderStyle.THIN)
              cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(java.awt.Color.GRAY))
              cellStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.WHITE))
              cellStyle.setFillPattern(FillPatternType.DIAMONDS)

              val col = row.createCell(dayIndex)
              val dayOfWeekStr = dayOfWeek.toString
              col.setCellValue(dayOfWeekStr.take(1) + dayOfWeekStr.drop(1).toLowerCase)
              col.setCellStyle(cellStyle)
            }

            sheet.addMergedRegion(new CellRangeAddress(3, 20, dayIndex, dayIndex))
          }
        }
      }
    }

    def resizeCols(implicit workbook: XSSFWorkbook) = {
      val sheet = workbook.getSheet(monthStr)

      for (i <- EmployeesInfoHeader.indices)
        sheet.autoSizeColumn(i)

      for (i <- 5 to (5 + numOfDays))
        sheet.setColumnWidth(i, 1200)
    }

    using(new XSSFWorkbook) { implicit workbook =>
      workbook.createSheet(monthStr)

      writeSheetHeader
      writeCompanyDetails
      writeAttendanceHeader
      writeAttendance
      writeHolidays
      writeWeekends

      resizeCols

      using(new FileOutputStream(AttendanceReportFileName))(workbook.write)
    }
  }
}
