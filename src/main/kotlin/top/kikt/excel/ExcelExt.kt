@file:Suppress("unused")

package top.kikt.excel

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.io.IOException

@Deprecated("Use letterToColumnIndex instead", ReplaceWith("letterToColumnIndex(letter)"))
fun letterToRowIndex(letter: String): Int {
    return letterToColumnIndex(letter)
}

/**
 * Example: "A" -> 0, "B" -> 1, "AA" -> 26, "AB" -> 27
 */
fun letterToColumnIndex(letter: String): Int {
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

/**
 * Convert excel tag to cell address.
 *
 * Such as: "A1" => row 0, column 0
 *          "B2" => row 1, column 1
 *          "AA1" => row 0, column 26
 */
fun String.toAddress(): CellAddress {
    val letter = this.replace(Regex("[0-9]"), "")
    val number = this.replace(Regex("[A-Z]"), "")

    val rowIndex = number.toInt() - 1
    val columnIndex = letterToColumnIndex(letter)

    return CellAddress(rowIndex, columnIndex)
}

/**
 * Get cell with row and column index.
 * Return null if not exist.
 */
fun Sheet.getCellOrNull(rowIndex: Int, columnIndex: Int): Cell? {
    val row = getRow(rowIndex) ?: return null
    return row.getCell(columnIndex)
}

/**
 * Get cell for CellAddress.
 */
fun Sheet.getCellOrNull(address: CellAddress): Cell? {
    return getCellOrNull(address.row, address.column)
}

/**
 * Get cell with "tag"
 * Return null if not found.
 *
 * Such as: "A1" => row 0, column 0
 *          "B2" => row 1, column 1
 *          "AA1" => row 0, column 26
 */
fun Sheet.getCellOrNull(tag: String): Cell? {
    val address = tag.toAddress()
    return getCellOrNull(address)
}

/**
 * Get cell with "tag"
 * If not found, create it.
 */
fun Sheet.getCellOrCreate(tag: String): Cell {
    val address = tag.toAddress()
    return getCellOrCreate(address)
}

/**
 * Get cell with address
 */
fun Sheet.getCellOrCreate(address: CellAddress): Cell {
    return getCellOrCreate(address.row, address.column)
}

/**
 * Get cell with row and column index.
 * If not found, create it.
 */
fun Sheet.getCellOrCreate(rowIndex: Int, columnIndex: Int): Cell {
    val row = getRow(rowIndex) ?: createRow(rowIndex)
    return row.getCell(columnIndex) ?: row.createCell(columnIndex)
}

/** Get Cell, maybe null */
fun Row.getCellOrNull(letter: String): Cell? {
    val index = letterToColumnIndex(letter)
    return getCell(index)
}

/**
 * Get Cell, if not exist, create it.
 */
fun Row.getCellOrCreate(letter: String): Cell {
    val index = letterToColumnIndex(letter)
    return getCell(index) ?: createCell(index)
}

/**
 * Create new cell, and set cell style.
 */
fun Row.createCell(letter: String, style: CellStyle? = null): Cell {
    val index = letterToColumnIndex(letter)
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

/**
 * Create workbook from file path.
 *
 * If deleteOld is true, delete old file.
 * otherwise, call the [File.toWorkbook] method to get workbook.
 */
fun File.createWorkbook(deleteOld: Boolean = false): Workbook {
    if (!deleteOld && exists()) {
        return toWorkbook()
    }
    when {
        extension.contentEquals("xls", true) -> {
            return HSSFWorkbook()
        }

        extension.contentEquals("xlsx", true) -> {
            return XSSFWorkbook()
        }
    }

    throw IllegalStateException("The file cannot create excel file.")
}