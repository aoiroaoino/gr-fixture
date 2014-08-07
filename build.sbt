import AssemblyKeys._

name := "gr-fixture"

organization := "jp.co.applicative.tool"

version := "0.1.0-SNAPSHOT"

scalaVersion := "2.10.4"

libraryDependencies ++= Seq(
  "org.scala-lang" % "scala-swing" % "2.10.2",
  "org.specs2" %% "specs2" % "1.13" % "test",
  "org.yaml" % "snakeyaml" % "1.12",
  "org.apache.poi" % "poi" % "3.9"
)

initialCommands := "import jp.co.applicative.tool._"

assemblySettings
