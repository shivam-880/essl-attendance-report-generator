package com.codingkapoor.esslattendancereportgenerator.model

import java.time.LocalDate

import com.codingkapoor.esslattendancereportgenerator.`package`._
import com.codingkapoor.esslattendancereportgenerator.core.RuntimeEnvironment
import play.api.libs.json.{Json, Reads}

import scala.io.Source

case class Holiday(date: LocalDate, occasion: String)

object Holiday {
  private lazy val holidaysAsJson: String = using(Source.fromFile(s"${RuntimeEnvironment.getDataDir}/$HolidaysFileName"))(_.mkString)

  private implicit val holidaysReads: Reads[Holiday] = Json.reads[Holiday]

  def getHolidays(month: Int, year: Int): List[Holiday] = Json.parse(holidaysAsJson).as[List[Holiday]].filter { holiday =>
    holiday.date.getMonthValue == month && holiday.date.getYear == year
  }
}
