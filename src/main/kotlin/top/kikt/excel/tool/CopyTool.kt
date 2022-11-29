package top.kikt.excel.tool

import org.apache.poi.ss.usermodel.*
import org.slf4j.LoggerFactory
import top.kikt.excel.isMerged
import top.kikt.excel.isMergedMainCell
import top.kikt.excel.stringValue

internal class CopyTool(val src: Sheet, val targetWorkbook: Workbook) {

    companion object {
        private val logger = LoggerFactory.getLogger(CopyTool::class.java)
    }

    fun copy(): Sheet {
        logger.trace("Start copy sheet ${src.sheetName}")

        val target = targetWorkbook.createSheet()

        /** The map key is src font index, value is target workbook font */
        val fontMap = mutableMapOf<Int, Font>()

        src.workbook.numberOfFonts.apply {
            for (i in 0 until this) {
                val font = src.workbook.getFontAt(i)

                targetWorkbook.createFont().apply {
                    fontName = font.fontName
                    bold = font.bold
                    italic = font.italic
                    strikeout = font.strikeout
                    typeOffset = font.typeOffset
                    underline = font.underline
                    color = font.color

                    charSet = font.charSet
                    fontHeight = font.fontHeight
                    fontHeightInPoints = font.fontHeightInPoints

                    fontMap[i] = this

                    logger.trace("src: {}", font)
                    logger.trace("target: {}", this)
                }
            }
        }

        src.getRow(0).getCell(0).showStyle()

        // copy merged region
        for (i in 0 until src.numMergedRegions) {
            val region = src.getMergedRegion(i)
            target.addMergedRegion(region)
        }

        for (row in src) {
            val targetRow = target.createRow(row.rowNum)
            // copy rowStyle
            if (row.rowStyle != null && targetRow.rowStyle == null) {
                targetRow.rowStyle.cloneStyleFrom(row.rowStyle)
            }

            for (cell in row) {
                val targetCell = targetRow.createCell(cell.columnIndex)
                cell.copyTo(targetCell)

                logger.debug("target cell style in row for each after copy to: {}", targetCell.cellStyle.debugInfo())

                logger.trace(
                    "row: {}, col: {}, foreground color: {}",
                    row.rowNum,
                    cell.columnIndex,
                    cell.cellStyle.fillForegroundColor
                )

                logger.debug("target cell style in row for each after set font: {}", targetCell.cellStyle.debugInfo())
            }
        }

        // evaluate all formula
        targetWorkbook.creationHelper.createFormulaEvaluator().evaluateAll()

        logger.trace("End copy sheet ${src.sheetName}")

        return target
    }


    /**
     *
     */
    internal fun Cell.showStyle() {
        // font style
        val index = cellStyle.fontIndex
        val font = row.sheet.workbook.getFontAt(index)
        logger.trace("index: {}, font: {}", index, font)
    }

    /**
     * Copy cell to target cell
     */
    fun Cell.copyTo(other: Cell) {
        if (isMerged() && !isMergedMainCell()) {
            return
        }
        logger.trace("copy cell before: {}, {}, src cell: {}", row.rowNum, columnIndex, this)
        run {
            // clone style
            val src = this.cellStyle
            val target = other.createStyle()

            if (src != null) {
                if (src.javaClass != target.javaClass) {
                    logger.debug("The style class is not same, src: {}, target: {}", src.javaClass, target.javaClass)
                    // use custom clone method
                    customCopyStyleTo(other)
                } else {
                    target.cloneStyleFrom(src)
                }
            }
        }

        other.cellComment = this.cellComment
        other.hyperlink = this.hyperlink

        try {
            when (cellType) {
                CellType.BLANK -> other.setBlank()
                CellType.BOOLEAN -> other.setCellValue(this.booleanCellValue)
                CellType.ERROR -> other.setCellErrorValue(this.errorCellValue)
                CellType.FORMULA -> other.cellFormula = this.cellFormula
                CellType.NUMERIC -> other.setCellValue(this.numericCellValue)
                CellType.STRING -> other.setCellValue(this.richStringCellValue)
                else -> {
                }
            }
        } catch (e: Exception) {
            logger.warn("set cell error", e)
            other.setCellValue(stringValue())
        }

        logger.trace("copy cell after: {}, {}, target cell: {}", row.rowNum, columnIndex, other)

        logger.debug("target cell style: {}", other.cellStyle.debugInfo())
    }

    /**
     * Custom copy cell style to target cell.
     */
    private fun Cell.customCopyStyleTo(other: Cell) {
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
            otherStyle.fillBackgroundColor = this.fillBackgroundColor
            otherStyle.fillForegroundColor = this.fillForegroundColor

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

    private fun Cell.createStyle(): CellStyle {
        val style = row.sheet.workbook.createCellStyle()
        cellStyle = style
        return style
    }

    private fun CellStyle.debugInfo(): String {
        val sb = StringBuilder()
        sb.append("fontIndex: $fontIndex")
        sb.append(", fillForegroundColor: $fillForegroundColor")
        sb.append(", fillBackgroundColor: $fillBackgroundColor")
        sb.append(", dataFormat: $dataFormat")
        sb.append(", alignment: $alignment")
        sb.append(", verticalAlignment: $verticalAlignment")
        sb.append(", borderBottom: $borderBottom")
        sb.append(", borderLeft: $borderLeft")
        sb.append(", borderRight: $borderRight")
        sb.append(", borderTop: $borderTop")
        sb.append(", bottomBorderColor: $bottomBorderColor")
        sb.append(", leftBorderColor: $leftBorderColor")
        sb.append(", rightBorderColor: $rightBorderColor")
        sb.append(", topBorderColor: $topBorderColor")
        sb.append(", wrapText: $wrapText")
        sb.append(", rotation: $rotation")
        sb.append(", indention: $indention")
        sb.append(", shrinkToFit: $shrinkToFit")
        sb.append(", hidden: $hidden")
        sb.append(", locked: $locked")
        return sb.toString()
    }

}