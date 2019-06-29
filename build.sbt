
name := "essl-attendance-report-generator"

version := "0.1"

scalaVersion := "2.12.8"

enablePlugins(JavaAppPackaging, UniversalPlugin)

libraryDependencies ++= Seq(
  "com.typesafe.scala-logging" %% "scala-logging" % "3.9.2",
  "ch.qos.logback" % "logback-classic" % "1.1.2",
  "com.typesafe" % "config" % "1.3.1"
)
