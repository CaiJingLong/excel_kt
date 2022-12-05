package top.kikt.excel

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Color
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * Check the cell is merged.
 */
fun Cell.isMerged(): Boolean {
    val sheet = row.sheet
    for (i in 0 until sheet.numMergedRegions) {
        val region = sheet.getMergedRegion(i)
        if (region.isInRange(row.rowNum, this.columnIndex)) {
            return true
        }
    }
    return false
}

/**
 * Check the cell is merged, and it is first cell of the merged region.
 */
fun Cell.isMergedMainCell(): Boolean {
    val sheet = row.sheet
    for (i in 0 until sheet.numMergedRegions) {
        val region = sheet.getMergedRegion(i)
        if (region.isInRange(row.rowNum, this.columnIndex)) {
            return region.firstRow == row.rowNum && region.firstColumn == this.columnIndex
        }
    }
    return false
}


/**
 * Copy cell font to target cell
 */
fun Cell.copyCellFont(other: Cell) {
    val srcWorkbook = this.workbook
    val targetWorkbook = other.workbook

    val srcFont = srcWorkbook.getFontAt(this.cellStyle.fontIndex)
    val targetFont = srcFont.copy(srcWorkbook, targetWorkbook)

    other.cellStyle.setFont(targetFont)
}

/**
 * The color to another workbook.
 */
fun Color.toColor(dstWorkbook: Workbook): Color? {
    val rgbArray = getColorRgb()

    if (rgbArray.count() != 3) {
        return null
    }

    var color: Color? = null
    if (dstWorkbook is XSSFWorkbook) {
        val xssf = XSSFColor(rgbArray)
        color = XSSFColor.from(xssf.ctColor, dstWorkbook.stylesSource.indexedColors)
    } else if (dstWorkbook is HSSFWorkbook) {
        color = dstWorkbook.customPalette.findSimilarColor(rgbArray[0], rgbArray[1], rgbArray[2])
    }

    return color
}

/**
 * Get the color rgb byte array.
 * The array length is 3.
 * The array element is 0-255.
 */
fun Color.getColorRgb(): ByteArray {
    val rgbArray = ArrayList<Byte>()
    if (this is XSSFColor) {
        argbHex?.let {
            val rgb = it.substring(2)
            for (i in rgb.indices step 2) {
                val hex = rgb.substring(i, i + 2)
                rgbArray.add(hex.toInt(16).toByte())
            }
        }
    } else if (this is HSSFColor) {
        triplet.forEach {
            rgbArray.add(it.toByte())
        }
    }

    return rgbArray.toByteArray()
}