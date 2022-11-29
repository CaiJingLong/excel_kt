package top.kikt.excel

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.usermodel.CellType.*
import org.slf4j.LoggerFactory
import top.kikt.excel.tool.CopyTool

private val logger = LoggerFactory.getLogger("SheetExt")

/**
 * Copy sheet from workbook to another workbook
 */
fun Sheet.copyTo(targetWorkbook: Workbook): Sheet {
   return CopyTool(this, targetWorkbook).copy()
}
