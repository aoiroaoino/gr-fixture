package jp.co.applicative.tool

import java.io.StringWriter
import java.util.LinkedHashMap
import java.io.StringWriter
import org.apache.poi.hssf.usermodel._
import org.apache.poi.ss.usermodel.{Cell, Row}
import scala.collection.JavaConversions._

object SQLHelper {

  def sheetProc(sheet: HSSFSheet, sheetName: String): String = {

    // 先読みの為Bufferedwriterへ
    val rowIter: BufferedIterator[Row] = sheet.rowIterator.buffered

    val header: List[String] =
      rowIter.head.cellIterator.map(getCellValue).toList

    val escapeCellIndex: List[Int] =
      header.zipWithIndex.filter{case (s, _) => s endsWith "$"}.map(_._2)

    val data: List[List[String]] =
      rowIter.drop(1).map(_.cellIterator.map(getCellValue).toList).toList
      // (for {
      //   itr <- rowIter.drop(1) // Header行を除く
      //   cell <- itr.cellIterator
      // } yield getCellValue(cell)).toList

    val dataLine = data map {
      _.zipWithIndex.map {
        case ((d: String, i: Int)) => if (escapeCellIndex.contains(i)) escape(d) else d
      }
    }

    val _dataLine = replaceLastCommaByColon(("" /: dataLine.map(dataTemplate))(_ + _))

    if (_dataLine.isEmpty) "" else headerTemplate(header map (_.replaceAll("$", ""))) + "\n" + _dataLine
  }

  private def headerTemplate(s: List[String]): String = s"""insert into (${s.mkString(",")}) values"""

  private def dataTemplate(s: List[String]): String = s"""(${s.mkString(",")}),\n"""

  private def replaceLastCommaByColon(s: String) = s.reverse.replaceFirst(",", ";").reverse

  private def escape(s: String): String = """"""" + s + """""""

  private def getCellValue(cell: Cell): String = PoiHelper.getCellValue(cell.asInstanceOf[HSSFCell])
}
