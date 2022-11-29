package top.kikt.excel

import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Row


/**
 * Fill color to cell style.
 */
fun Row.fillColor(
    color: HSSFColor.HSSFColorPredefined = HSSFColor.HSSFColorPredefined.LIGHT_YELLOW,
    fillBorder: Boolean = true,
    borderColor: HSSFColor.HSSFColorPredefined = HSSFColor.HSSFColorPredefined.BLACK,
    borderStyle: BorderStyle = BorderStyle.THIN,
) {
    val style = workbook.createCellStyle()
    style.fillForegroundColor = color.index
    style.fillPattern = FillPatternType.SOLID_FOREGROUND
    if (fillBorder) {
        style.makeBorder(borderColor = borderColor, borderStyle = borderStyle)
    }
    for (cell in this) {
        cell.cellStyle = style
    }
}

/**
 * Make border to cell style.
 */
fun CellStyle.makeBorder(
    borderColor: HSSFColor.HSSFColorPredefined = HSSFColor.HSSFColorPredefined.BLACK,
    borderStyle: BorderStyle = BorderStyle.THIN
) {
    borderLeft = borderStyle
    borderTop = borderStyle
    borderRight = borderStyle
    borderBottom = borderStyle
    this.leftBorderColor = borderColor.index
    this.topBorderColor = borderColor.index
    this.rightBorderColor = borderColor.index
    this.bottomBorderColor = borderColor.index
}
