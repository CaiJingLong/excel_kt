package top.kikt.excel

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import top.kikt.excel.tool.CopySheetTool

/**
 * Copy sheet from workbook to another workbook
 */
fun Sheet.copyTo(
    targetWorkbook: Workbook,
    index: Int? = null,
    name: String? = null,
    active: Boolean = false,
): Sheet {
    return CopySheetTool(this, targetWorkbook).copy(index, name, active)
}
