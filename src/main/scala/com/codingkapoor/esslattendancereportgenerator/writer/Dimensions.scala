package com.codingkapoor.esslattendancereportgenerator.writer

sealed trait Dimensions

case class CompanyDetailsDimensions(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions

case class SectionHeaderDimensions(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions

case class EmployeeInfoHeaderDimensions(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions

case class EmployeeInfoDimensions(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions

case class AttendanceHeaderDimensions(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions

case class AttendanceDimensions(firstRowIndex: Int, lastRowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int) extends Dimensions
