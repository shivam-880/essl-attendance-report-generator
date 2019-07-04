package com.codingkapoor.esslattendancereportgenerator

object `package` {
  val CompanyName = "Glassbeam Software India Pvt Ltd"
  val CompanyAddress = "No. 21, 3rd Floor, Block A, Sree Rama Deevana, Halasuru Road, Bangalore-42"

  val EmployeesInfoHeader = Array("Employee Name", "Gender", "Date Of Joining", "PF No.")

  val AttendanceReportFileName = "gbatt.xlsx"

  def using[A <: {def close() : Unit}, B](resource: A)(f: A => B): B = {
    try {
      f(resource)
    } finally {
      resource.close()
    }
  }
}
