package com.codingkapoor.esslattendancereportgenerator.model

import java.time.LocalDate

import com.codingkapoor.esslattendancereportgenerator.`package`._
import com.codingkapoor.esslattendancereportgenerator.core.RuntimeEnvironment
import play.api.libs.json.{Json, Reads}

import scala.collection.immutable
import scala.io.Source

case class Employee(empId: Int, name: String, gender: String, doj: LocalDate, pfn: String)

object Employee {
  private lazy val employeesAsJson: String = using(Source.fromFile(s"${RuntimeEnvironment.getDataDir}/$EmployeesFileName"))(_.mkString)

  private implicit val employeesReads: Reads[Employee] = Json.reads[Employee]

  def getEmployees: immutable.Seq[Employee] = Json.parse(employeesAsJson).as[List[Employee]]
}
