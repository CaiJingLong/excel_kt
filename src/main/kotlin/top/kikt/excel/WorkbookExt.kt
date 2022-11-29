package top.kikt.excel

import org.apache.poi.ss.usermodel.*
import org.slf4j.LoggerFactory
import top.kikt.excel.tool.CopySheetTool

private val logger = LoggerFactory.getLogger("SheetExt")

/**
 * Copy sheet from workbook to another workbook
 */
fun Sheet.copyTo(targetWorkbook: Workbook): Sheet {
   return CopySheetTool(this, targetWorkbook).copy()
}
