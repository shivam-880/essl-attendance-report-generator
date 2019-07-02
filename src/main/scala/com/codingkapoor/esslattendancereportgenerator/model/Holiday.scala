package com.codingkapoor.esslattendancereportgenerator.model

import java.time.LocalDate

import com.codingkapoor.esslattendancereportgenerator.`package`.using
import com.codingkapoor.esslattendancereportgenerator.core.RuntimeEnvironment
import play.api.libs.json.{Json, Reads}

import scala.collection.immutable
import scala.io.Source

case class Holiday(date: LocalDate, occasion: String)

object Holiday {
  private lazy val holidaysAsJson: String = using(Source.fromFile(s"${RuntimeEnvironment.getDataDir}/holidays.json"))(_.mkString)

  private implicit val holidaysReads: Reads[Holiday] = Json.reads[Holiday]

  def getHolidays: immutable.Seq[Holiday] = Json.parse(holidaysAsJson).as[List[Holiday]]
}
