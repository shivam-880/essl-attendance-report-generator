package com.codingkapoor.esslattendancereportgenerator

object `package` {
  val CompanyName = "GLASSBEAM SOFTWARE INDIA PVT. LTD.\n"
  val CompanyAddress = "No. 21, 3RD FLOOR, BLOCK A, SREE RAMA DEEVANA,\nHALASURU ROAD, BANGALORE-42"

  val EmployeesInfoHeader = Array("ID", "EMPLOYEE NAME", "GENDER", "DATE OF JOINING", "PF NUMBER")

  val AttendanceReportFileName = "gbatt.xlsx"

  object AttendanceStatus extends Enumeration {
    type AttendanceStatus = Value

    val Present: AttendanceStatus.Value = Value("P")
    val Leave: AttendanceStatus.Value = Value("L")
    val Holiday: AttendanceStatus.Value = Value("H")
    val HalfDay: AttendanceStatus.Value = Value("1/2")
    val Wfh: AttendanceStatus.Value = Value("WFH")
  }

  def using[A <: {def close() : Unit}, B](resource: A)(f: A => B): B = {
    try {
      f(resource)
    } finally {
      resource.close()
    }
  }
}
