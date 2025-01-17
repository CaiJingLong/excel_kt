package top.kikt.excel

import org.apache.poi.hssf.usermodel.HSSFRichTextString
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFRichTextString
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
 * Get the main cell of the merged region.
 *
 * If the cell is not merged cell, return itself.
 */
fun Cell.getMainCell(): Cell {
    if (!isMerged()) {
        return this
    }

    val sheet = row.sheet
    val region = sheet.mergedRegions.find { it.isInRange(row.rowNum, 1) }
    if (region != null) {
        return sheet.getRow(region.firstRow).getCell(region.firstColumn)
    }

    // if not found, return itself，this is impossible
    return this
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
 * Copy cell style to target cell
 */
fun Cell.copyStyle(other: Cell) {
    val src = this.cellStyle
    val target = other.workbook.createCellStyle()

    if (src?.javaClass == target?.javaClass) {
        target.cloneStyleFrom(src)
        return
    }

    val otherStyle = other.cellStyle
    // ignore font , because font is alone method

    // copy border
    with(this.cellStyle) {
        otherStyle.borderBottom = this.borderBottom
        otherStyle.borderLeft = this.borderLeft
        otherStyle.borderRight = this.borderRight
        otherStyle.borderTop = this.borderTop

        // copy border color
        otherStyle.bottomBorderColor = this.bottomBorderColor
        otherStyle.leftBorderColor = this.leftBorderColor
        otherStyle.rightBorderColor = this.rightBorderColor
        otherStyle.topBorderColor = this.topBorderColor

        // copy fill
        this.fillBackgroundColorColor?.toColor(other.workbook)?.toColorIndex(other.workbook)
            ?.let {
                otherStyle.fillBackgroundColor = it
            }
        this.fillForegroundColorColor?.toColor(other.workbook)?.toColorIndex(other.workbook)
            ?.let {
                otherStyle.fillForegroundColor = it
            }

        // copy fill pattern
        otherStyle.fillPattern = this.fillPattern

        // copy alignment
        otherStyle.alignment = this.alignment
        otherStyle.verticalAlignment = this.verticalAlignment

        // copy wrap text
        otherStyle.wrapText = this.wrapText

        // copy date format
        otherStyle.dataFormat
    }
}

private fun Color.toColorIndex(dst: Workbook): Short? {
    val color = toColor(dst)
    if (color is XSSFColor) {
        return color.indexed
    } else if (color is HSSFColor) {
        return color.index
    }
    return null
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

/**
 * Copy cell value to target cell.
 */
fun Cell.copyCellValue(other: Cell) {
    try {
        when (cellType) {
            CellType.BLANK -> other.setBlank()
            CellType.BOOLEAN -> other.setCellValue(this.booleanCellValue)
            CellType.ERROR -> other.setCellErrorValue(this.errorCellValue)
            CellType.FORMULA -> other.cellFormula = this.cellFormula
            CellType.NUMERIC -> other.setCellValue(this.numericCellValue)
            CellType.STRING -> this.copyStringValueTo(other)
            else -> {
            }
        }
    } catch (e: Exception) {
        other.setCellValue(stringValue())
    }
}

/**
 * The cell is same class with other cell.
 */
fun Cell.isSameClass(other: Cell): Boolean {
    return this.javaClass == other.javaClass
}

private fun Cell.copyStringValueTo(other: Cell) {
    if (cellType != CellType.STRING) {
        return
    }
    if (isSameClass(other)) {
        other.setCellValue(this.richStringCellValue)
        return
    }

    // if workbook is not same type, use string value
    val targetValue: RichTextString =
        when (val richStringCellValue = this.richStringCellValue) {
            is XSSFRichTextString -> {
                convertToHSSFRichTextString(richStringCellValue, other.workbook)
            }

            is HSSFRichTextString -> {
                convertToXSSFRichTextString(richStringCellValue, other.workbook)
            }

            else -> {
                null
            }
        } ?: return

    other.setCellValue(targetValue)
}

private fun Cell.convertToHSSFRichTextString(src: XSSFRichTextString, targetWorkbook: Workbook): HSSFRichTextString {
    val hssfRichTextString = HSSFRichTextString(src.string)

    for (formatIndex in 0 until src.numFormattingRuns()) {
        val srcFont = src.getFontOfFormattingRun(formatIndex) ?: continue
        val formatStart = src.getIndexOfFormattingRun(formatIndex)
        val formatLength = src.getLengthOfFormattingRun(formatIndex)
        val targetFont = srcFont.copy(workbook, targetWorkbook)

        hssfRichTextString.applyFont(formatStart, formatStart + formatLength, targetFont)
    }

    return hssfRichTextString
}

private fun Cell.convertToXSSFRichTextString(src: HSSFRichTextString, targetWorkbook: Workbook): XSSFRichTextString {
    val xssfRichTextString = XSSFRichTextString(src.string)

    for (i in 0 until src.string.count()) {
        val fontIndex = src.getFontAtIndex(i)
        val font = workbook.getFontAt(fontIndex.toInt())
        val targetFont = font.copy(workbook, targetWorkbook)
        xssfRichTextString.applyFont(i, i + 1, targetFont)
    }

    return xssfRichTextString
}
