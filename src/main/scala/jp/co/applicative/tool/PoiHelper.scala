package jp.co.applicative.tool

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.DateUtil

object PoiHelper {

  def getCellValue(cell: HSSFCell): String = {
    if (cell == null) return ""
    cell.getCellType() match {
      case Cell.CELL_TYPE_BLANK => ""
      case Cell.CELL_TYPE_STRING => cell.getStringCellValue()
      case Cell.CELL_TYPE_BOOLEAN => cell.getBooleanCellValue()
      case Cell.CELL_TYPE_NUMERIC => {
        cell match {
          case c if DateUtil.isCellDateFormatted(c) => c.getDateCellValue()
          case c => c.getNumericCellValue() match {
            case v if v == v.intValue => v.intValue.toString // Double にされる前に文字列化
            case v => v
          }
        }
      }
      case Cell.CELL_TYPE_FORMULA => {
        val fe = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator()
        getCellValue(fe.evaluateInCell(cell))
      }
      case _ => ""
    }
  }.toString
}
