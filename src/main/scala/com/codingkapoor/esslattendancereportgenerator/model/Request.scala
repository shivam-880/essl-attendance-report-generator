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

  def getRequests: Map[Int, Map[LocalDate, String]] =
    Json.parse(requestsAsJson).as[List[Request]].foldLeft(Map[Int, Map[LocalDate, String]]()) { (acc, request) =>
      val i = Map(request.date -> request.req)
      if (acc.contains(request.empId))
        acc + (request.empId -> (acc(request.empId) ++: i))
      else
        acc + (request.empId -> i)
    }
}
