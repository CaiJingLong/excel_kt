package top.kikt.excel

import org.apache.poi.ss.usermodel.Cell

/**
 *
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
