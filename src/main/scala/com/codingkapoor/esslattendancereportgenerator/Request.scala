package com.codingkapoor.esslattendancereportgenerator

import java.time.LocalDate

import play.api.libs.json.{Json, Reads}

import scala.collection.immutable
import scala.io.Source

case class Request(empId: Int, req: String, date: LocalDate)

object Request {
  private lazy val requestsAsJson: String = using(Source.fromFile(s"${RuntimeEnvironment.getDataDir}/requests.json"))(_.mkString)

  private implicit val employeesReads: Reads[Request] = Json.reads[Request]

  def getRequests: immutable.Seq[Request] = Json.parse(requestsAsJson).as[List[Request]]
}
