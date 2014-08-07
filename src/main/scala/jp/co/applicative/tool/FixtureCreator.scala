package jp.co.applicative.tool

import java.io._
import org.apache.poi.hssf.usermodel._
import org.apache.poi.ss.usermodel.Cell
import scala.collection.mutable.LinkedHashMap
import scala.collection.JavaConversions._

object FixtureCreator {

  val DOUBLEQUOTE = """""""

  def create(inPath: String, outPath: String): Unit = {

    //Excel Load
    val wb = new HSSFWorkbook(new FileInputStream(inPath))

    //Sheet Loop
    for (i <- 0 until wb.getNumberOfSheets()) {
      val sheetName = wb.getSheetName(i)
      writeFixture(sheetProc(wb.getSheetAt(i), sheetName), outPath, sheetName + ".yml")
    }
  }

  def sheetProc(sheet: HSSFSheet, sheetName: String): String = {
    val map = LinkedHashMap.empty[String, Any]
    val headerRow = sheet.getRow(0)
    val itr = sheet.rowIterator()

    //Row Loop
    while (itr.hasNext()) {
      val row = itr.next().asInstanceOf[HSSFRow]
      row.getRowNum() match {
        case 0 => //skip header row
        case _ => if (PoiHelper.getCellValue(row.getCell(0)) != "") {
          map += (f"${sheetName}_${row.getRowNum()}" -> rowProc(headerRow, row))
        }
      }
    }

    YamlHelper.dump(convertJavaMap(map))
  }

  def rowProc(header: HSSFRow, row: HSSFRow): LinkedHashMap[String, String] = {
    val map = LinkedHashMap.empty[String, String]
    val itr = header.cellIterator()
    while (itr.hasNext()) {
      val key = itr.next().asInstanceOf[HSSFCell]
      val value = PoiHelper.getCellValue(row.getCell(key.getCellNum()))
      PoiHelper.getCellValue(key) match {
        case k if (!k.endsWith("$")) => map += (k -> value)
        case k if (value.lines.length > 1) => map += (k.substring(0, k.length() - 1) -> value)
        case k => map += (k.substring(0, k.length() - 1) -> escape(value))
      }
    }
    return map
  }

  def escape(target: Any): String = DOUBLEQUOTE + target + DOUBLEQUOTE

  def convertJavaMap(map: LinkedHashMap[String, Any]): java.util.LinkedHashMap[String, Object] = {
    val javaMap: java.util.LinkedHashMap[String, Object] = new java.util.LinkedHashMap
    map.foreach {
      case (key, value) => {
        javaMap.put(key.toString, value match {
          case v: LinkedHashMap[String, Any] => convertJavaMap(v)
          case _ => value.asInstanceOf[Object]
        })
      }
    }
    return javaMap
  }

  def writeFixture(data: String, path: String, fileName: String) = {
    if (data.isEmpty() == false) {
      val writer = new PrintWriter(new BufferedWriter(new OutputStreamWriter(new FileOutputStream(new File(path, fileName), false), "UTF-8")))
      writer.write(data)
      writer.close()
    }
  }

}
