package com.codingkapoor.esslattendancereportgenerator.writer

import java.time.YearMonth

import com.codingkapoor.esslattendancereportgenerator.`package`.EmployeesInfoHeader
import com.codingkapoor.esslattendancereportgenerator.model.Employee

sealed trait Dimensions

case class CompanyDetailsDimensions private(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions

object CompanyDetailsDimensions {
  def apply(): CompanyDetailsDimensions = {
    new CompanyDetailsDimensions(
      firstRowIndex = 0,
      lastRowIndex = 0,
      firstColumnIndex = 0,
      lastColumnIndex = EmployeesInfoHeader.length - 1)
  }
}

case class SectionHeaderDimensions private(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions

object SectionHeaderDimensions {
  def apply(month: Int, year: Int): SectionHeaderDimensions = {
    val yearMonth = YearMonth.of(year, month)
    val numOfDaysInMonth = yearMonth.lengthOfMonth

    new SectionHeaderDimensions(
      firstRowIndex = 0,
      lastRowIndex = 0,
      firstColumnIndex = EmployeesInfoHeader.length,
      lastColumnIndex = numOfDaysInMonth + EmployeesInfoHeader.length - 1
    )
  }
}

case class EmployeeInfoHeaderDimensions private(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions

object EmployeeInfoHeaderDimensions {
  def apply(): EmployeeInfoHeaderDimensions = {
    new EmployeeInfoHeaderDimensions(
      firstRowIndex = 2,
      lastRowIndex = 2,
      firstColumnIndex = 0,
      lastColumnIndex = EmployeesInfoHeader.length - 1
    )
  }
}

case class EmployeeInfoDimensions private(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions

object EmployeeInfoDimensions {
  def apply(employees: Seq[Employee]): EmployeeInfoDimensions = {

    new EmployeeInfoDimensions(
      firstRowIndex = 3,
      lastRowIndex = employees.size + 3,
      firstColumnIndex = 0,
      lastColumnIndex = EmployeesInfoHeader.length - 1
    )
  }
}

case class AttendanceHeaderDimensions private(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions

object AttendanceHeaderDimensions {
  def apply(month: Int, year: Int): AttendanceHeaderDimensions = {
    val yearMonth = YearMonth.of(year, month)
    val numOfDaysInMonth = yearMonth.lengthOfMonth

    new AttendanceHeaderDimensions(
      firstRowIndex = 2,
      lastRowIndex = 2,
      firstColumnIndex = EmployeesInfoHeader.length,
      lastColumnIndex = numOfDaysInMonth + EmployeesInfoHeader.length - 1
    )
  }
}

case class AttendanceDimensions private(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions

object AttendanceDimensions {
  def apply(month: Int, year: Int, employees: Seq[Employee]): AttendanceDimensions = {
    val yearMonth = YearMonth.of(year, month)
    val numOfDaysInMonth = yearMonth.lengthOfMonth

    new AttendanceDimensions(
      firstRowIndex = 3,
      lastRowIndex = employees.size + 3,
      firstColumnIndex = EmployeesInfoHeader.length,
      lastColumnIndex = numOfDaysInMonth + EmployeesInfoHeader.length - 1
    )
  }
}
