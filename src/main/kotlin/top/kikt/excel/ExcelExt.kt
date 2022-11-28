@file:Suppress("unused")

package top.kikt.excel

import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.*
import java.io.FileOutputStream
import java.io.IOException


fun letterToRowIndex(letter: String): Int {
    letter.uppercase().apply {
        if (this.length == 1) {
            return letter[0] - 'A'
        }

        if (this.length == 2) {
            val firstIndex = this[0] - 'A'
            val secondIndex = this[1] - 'A'

            return firstIndex * 26 + secondIndex
        }

    }
    throw IOException("Invalid column letter")
}

/** 获取 Cell，可能为空 */
fun Row.getCellOrNull(letter: String): Cell? {
    val index = letterToRowIndex(letter)
    return getCell(index)
}

/** 获取 Cell，如果不存在，创建一个 */
fun Row.getCellOrCreate(letter: String): Cell {
    val index = letterToRowIndex(letter)
    return getCell(index) ?: createCell(index)
}

/** 创建 Cell */
fun Row.createCell(letter: String, style: CellStyle? = null): Cell {
    val index = letterToRowIndex(letter)
    return createCell(index).apply {
        if (style != null) {
            this.cellStyle = style
        }
    }
}


fun Cell.intValue(): Int {
    return try {
        when (cellType) {
            CellType.NUMERIC -> numericCellValue.toInt()
            CellType.STRING -> stringValue().toInt()
            else -> 0
        }
    } catch (e: Exception) {
        0
    }
}


fun Cell?.doubleValue(): Double {
    if (this == null) return 0.0
    return try {
        when (cellType) {
            CellType.NUMERIC -> numericCellValue
            CellType.STRING -> stringValue().toDouble()
            else -> 0.0
        }
    } catch (e: Exception) {
        0.0
    }
}

fun Cell?.stringValue(): String {
    if (this == null) return ""
    return try {
        when (cellType) {
            CellType.NUMERIC -> numericCellValue.toString()
            CellType.STRING -> stringCellValue
            else -> ""
        }
    } catch (e: Exception) {
        ""
    }
}

fun Row.getWorkbook(): Workbook {
    return sheet.workbook
}

fun Workbook.saveTo(outputPath: String) {
    outputPath.createFileIfExists()
    val fos = FileOutputStream(outputPath)
    this.write(fos)
    fos.close()
}

fun Row.fillColor(
    color: HSSFColor.HSSFColorPredefined = HSSFColor.HSSFColorPredefined.LIGHT_YELLOW,
    fillBorder: Boolean = true,
    borderColor: HSSFColor.HSSFColorPredefined = HSSFColor.HSSFColorPredefined.BLACK,
    borderStyle: BorderStyle = BorderStyle.THIN,
) {
    val wb = getWorkbook()
    val style = wb.createCellStyle()
    style.fillForegroundColor = color.index
    style.fillPattern = FillPatternType.SOLID_FOREGROUND
    if (fillBorder) {
        style.makeBorder(borderColor = borderColor, borderStyle = borderStyle)
    }
    for (cell in this) {
        cell.cellStyle = style
    }
}

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