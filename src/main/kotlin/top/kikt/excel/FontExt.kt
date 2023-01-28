package top.kikt.excel

import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.Color
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * Copy font from [srcWorkbook] to [targetWorkbook].
 */
fun Font.copy(srcWorkbook: Workbook, targetWorkbook: Workbook): Font {
    val result = targetWorkbook.createFont().apply {
        fontName = this@copy.fontName
        fontHeight = this@copy.fontHeight
        fontHeightInPoints = this@copy.fontHeightInPoints
        italic = this@copy.italic
        strikeout = this@copy.strikeout
        typeOffset = this@copy.typeOffset
        underline = this@copy.underline
        charSet = this@copy.charSet
        bold = this@copy.bold
    }

    var color: Color? = null

    if (this is XSSFFont) {
        color = xssfColor.toColor(targetWorkbook)

    } else if (this is HSSFFont) {
        val hssfColor = this.getHSSFColor(srcWorkbook as HSSFWorkbook)
        color = hssfColor.toColor(targetWorkbook)
    }

    if (color == null) {
        return result
    }

    if (targetWorkbook is XSSFWorkbook) {
        (result as XSSFFont).setColor(color as XSSFColor)
    } else if (targetWorkbook is HSSFWorkbook) {
        (result as HSSFFont).color = (color as HSSFColor).index
    }

    return result
}