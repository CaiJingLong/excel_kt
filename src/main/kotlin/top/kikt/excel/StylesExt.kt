package top.kikt.excel

import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellAddress


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

/**
 * Fill color for sheet range.
 */
fun Sheet.fillBorder(
    startAddress: CellAddress,
    endAddress: CellAddress,
    borderColor: HSSFColor.HSSFColorPredefined = HSSFColor.HSSFColorPredefined.BLACK,
    borderStyle: BorderStyle = BorderStyle.THIN
) {
    val style = workbook.createCellStyle()
    style.makeBorder(borderColor = borderColor, borderStyle = borderStyle)
    for (row in startAddress.row..endAddress.row) {
        val rowObj = getRow(row) ?: createRow(row)
        for (col in startAddress.column..endAddress.column) {
            val cell = rowObj.getCell(col) ?: rowObj.createCell(col)
            cell.cellStyle = style
        }
    }
}