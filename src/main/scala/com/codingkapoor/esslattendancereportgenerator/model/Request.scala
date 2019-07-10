package com.codingkapoor.esslattendancereportgenerator.model

import java.time.LocalDate

import com.codingkapoor.esslattendancereportgenerator.`package`._
import com.codingkapoor.esslattendancereportgenerator.core.RuntimeEnvironment
import play.api.libs.json.{Json, Reads}

import scala.io.Source

case class Request(empId: Int, req: String, date: LocalDate)

object Request {
  private lazy val requestsAsJson: String = using(Source.fromFile(s"${RuntimeEnvironment.getDataDir}/$RequestsFileName"))(_.mkString)

  private implicit val employeesReads: Reads[Request] = Json.reads[Request]

  def getRequests: List[Request] = Json.parse(requestsAsJson).as[List[Request]]
}
