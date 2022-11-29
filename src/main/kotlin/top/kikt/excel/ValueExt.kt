package top.kikt.excel

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType


/**
 * Get cell value, if cell is null, return 0.
 */
fun Cell?.intValue(): Int {
    return try {
        when (this?.cellType) {
            CellType.NUMERIC -> numericCellValue.toInt()
            CellType.STRING -> stringValue().toInt()
            else -> 0
        }
    } catch (e: Exception) {
        0
    }
}

/**
 * Get cell value, if cell is null, return 0.
 */
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

/**
 * Get cell value, if cell is null, return empty string.
 */
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
