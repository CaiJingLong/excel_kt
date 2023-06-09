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

/**
 * Get last not null row index.
 */
fun Sheet.getLastNotNullRowIndex(): Int {
    val row = lastRowNum
    for (i in row downTo 0) {
        val rowObj = getRow(i) ?: continue
        val cell = rowObj.lastCellNum
        for (j in cell downTo 0) {
            val cellObj = rowObj.getCell(j) ?: continue
            if (cellObj.stringValue().isNotBlank()) {
                return i
            }
        }
    }
    return 0
}

/**
 * Get last not null column index.
 */
fun Sheet.getLastNotNullColumnIndex(): Int {
    // not same row, the method need for each all row
    val rows = toList()

    if (rows.isEmpty()) {
        return -1
    }

    val values = ArrayList<Int>()

    for (row in rows) {
        for (cell in row.toList().asReversed()) {
            if (cell.stringValue().isNotBlank()) {
                values.add(cell.columnIndex)
                break
            }
        }
    }

    return values.maxOrNull() ?: -1
}

/**
 * Get first not null row index.
 */
fun Sheet.getFirstNotNullRowIndex(): Int {
    val row = firstRowNum
    for (i in row..lastRowNum) {
        val rowObj = getRow(i) ?: continue
        val cell = rowObj.firstCellNum
        for (j in cell..rowObj.lastCellNum) {
            val cellObj = rowObj.getCell(j) ?: continue
            if (cellObj.stringValue().isNotBlank()) {
                return i
            }
        }
    }
    return -1
}

/**
 * Get first not null column index.
 */
fun Sheet.getFirstNotNullColumnIndex(): Int {
    // not same row, the method need for each all row
    val rows = toList()

    if (rows.isEmpty()) {
        return -1
    }

    val values = ArrayList<Int>()

    for (row in rows) {
        for (cell in row.toList()) {
            if (cell.stringValue().isNotBlank()) {
                values.add(cell.columnIndex)
                break
            }
        }
    }

    return values.minOrNull() ?: -1
}