@file:Suppress("unused")

package top.kikt.excel

import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.*
import java.io.File
import java.io.FileOutputStream
import java.io.IOException

/**
 * Example: "A" -> 0, "B" -> 1, "AA" -> 26, "AB" -> 27
 */
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

/** Get Cell, maybe null */
fun Row.getCellOrNull(letter: String): Cell? {
    val index = letterToRowIndex(letter)
    return getCell(index)
}

/**
 * Get Cell, if not exist, create it.
 */
fun Row.getCellOrCreate(letter: String): Cell {
    val index = letterToRowIndex(letter)
    return getCell(index) ?: createCell(index)
}

/**
 * Create new cell, and set cell style.
 */
fun Row.createCell(letter: String, style: CellStyle? = null): Cell {
    val index = letterToRowIndex(letter)
    return createCell(index).apply {
        if (style != null) {
            this.cellStyle = style
        }
    }
}

/**
 * Get workbook from row.
 */
val Row.workbook: Workbook
    get() = sheet.workbook

/**
 * Get workbook from cell.
 */
val Cell.workbook: Workbook
    get() = row.workbook

/**
 * Save workbook to file path.
 */
fun Workbook.saveTo(outputPath: String) {
    outputPath.createIfNotExists()
    FileOutputStream(outputPath).use {
        write(it)
    }
}

/**
 * Save workbook to file.
 */
fun Workbook.saveTo(file: File) {
    file.createIfNotExists()
    FileOutputStream(file).use {
        write(it)
    }
}

/**
 * File path to workbook.
 */
fun String.toWorkbook(): Workbook {
    return ExcelUtils.getWorkbook(this)
}

/**
 * File to workbook.
 */
fun File.toWorkbook(): Workbook {
    return ExcelUtils.getWorkbook(this)
}
