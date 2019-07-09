package com.codingkapoor.esslattendancereportgenerator.model

import java.time.LocalDate

import com.codingkapoor.esslattendancereportgenerator.`package`.using
import com.codingkapoor.esslattendancereportgenerator.core.RuntimeEnvironment
import play.api.libs.json.{Json, Reads}

import scala.collection.immutable
import scala.io.Source

case class Request(empId: Int, req: String, date: LocalDate)

object Request {
  private lazy val requestsAsJson: String = using(Source.fromFile(s"${RuntimeEnvironment.getDataDir}/requests.json"))(_.mkString)

  private implicit val employeesReads: Reads[Request] = Json.reads[Request]

  def getRequests: List[Request] = Json.parse(requestsAsJson).as[List[Request]]
}
