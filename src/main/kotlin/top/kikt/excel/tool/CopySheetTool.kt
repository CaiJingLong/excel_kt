package top.kikt.excel.tool

import org.apache.poi.hssf.usermodel.HSSFRichTextString
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.slf4j.LoggerFactory
import top.kikt.excel.*

internal class CopySheetTool(private val src: Sheet, private val targetWorkbook: Workbook) {

    companion object {
        private val logger = LoggerFactory.getLogger(CopySheetTool::class.java)
    }

    /** The map key is src font index, value is target workbook font */
    private val fontMap = mutableMapOf<Int, Font>()
    private val srcFontMap = mutableMapOf<Int, Font>()

    private val isSameType: Boolean by lazy {
        src.workbook.javaClass == targetWorkbook.javaClass
    }

    fun copy(index: Int? = null, targetName: String? = null, active: Boolean = false): Sheet {
        logger.trace("Start copy sheet ${src.sheetName}")

        val target = if (targetName == null) {
            targetWorkbook.createSheet()
        } else {
            targetWorkbook.createSheet(targetName)
        }

        if (index != null) {
            targetWorkbook.setSheetOrder(target.sheetName, index)
        }

        if (active) {
            val targetIndex = targetWorkbook.getSheetIndex(target)
            if (targetIndex != -1) {
                targetWorkbook.setActiveSheet(targetIndex)
            }
        }

        fontMap.clear()

        refreshFontMap()

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

                // copy cell width
//                val columnWidth = src.getColumnWidthInPixels(cell.columnIndex)
                target.setColumnWidth(cell.columnIndex, src.getColumnWidth(cell.columnIndex))
            }

            // copy height
            targetRow.heightInPoints = row.heightInPoints
        }

        // evaluate all formula
        targetWorkbook.creationHelper.createFormulaEvaluator().evaluateAll()

        logger.trace("End copy sheet ${src.sheetName}")

        return target
    }

    private fun Font.copyFont(): Font {
        val font = this
        return targetWorkbook.createFont().also {
            it.fontName = font.fontName
            it.bold = font.bold
            it.italic = font.italic
            it.strikeout = font.strikeout
            it.typeOffset = font.typeOffset
            it.underline = font.underline
            it.color = font.color

            it.charSet = font.charSet
            it.fontHeight = font.fontHeight
            it.fontHeightInPoints = font.fontHeightInPoints
        }
    }

    private fun refreshFontMap() {
        src.workbook.numberOfFonts.apply {
            for (i in 0 until this) {
                val font = src.workbook.getFontAt(i)

                srcFontMap[i] = font

                fontMap[i] = font.copyFont()
            }
        }
    }


    /**
     *
     */
    private fun Cell.showStyle() {
        // font style
        val index = cellStyle.fontIndex
        val font = row.sheet.workbook.getFontAt(index)
        logger.trace("index: {}, font: {}", index, font)
    }

    /**
     * Copy cell to target cell
     */
    private fun Cell.copyTo(other: Cell) {
        if (isMerged() && !isMergedMainCell()) {
            return
        }

        logger.trace("copy cell( {}, {} ) before: src cell: {}", row.rowNum, columnIndex, this)

        other.cellComment = this.cellComment
        other.hyperlink = this.hyperlink

        copyCellValue(other)
        copyCellStyle(other)
        copyCellFont(other)

        logger.trace("copy cell( {}, {} ) after : target cell: {}", row.rowNum, columnIndex, other)

        logger.debug("target cell style: {}", other.cellStyle.debugInfo())
    }

    private fun Cell.copyCellValue(other: Cell) {
        try {
            logger.trace("copy cell({}, {}) value before, src cell: {}", row.rowNum, columnIndex, this)
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
            logger.trace("copy cell({}, {}) value after, src cell: {}", row.rowNum, columnIndex, this)
        } catch (e: Exception) {
            logger.warn("set cell error", e)
            other.setCellValue(stringValue())
        }
    }

    private fun Cell.copyCellStyle(other: Cell) {
        logger.debug("Copy cell style before: {}", other.cellStyle.debugInfo())
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

        logger.debug("Copy cell style after: {}", other.cellStyle.debugInfo())
    }

    private fun Cell.copyStringValueTo(other: Cell) {
        if (cellType != CellType.STRING) {
            return
        }
        if (isSameType) {
            other.setCellValue(this.richStringCellValue)
            return
        }

        // if workbook is not same type, use string value
        val targetValue: RichTextString =
            when (val richStringCellValue = this.richStringCellValue) {
                is XSSFRichTextString -> {
                    convertToHSSFRichTextString(richStringCellValue)
                }

                is HSSFRichTextString -> {
                    convertToXSSFRichTextString(richStringCellValue)
                }

                else -> {
                    null
                }
            } ?: return

        other.setCellValue(targetValue)
    }

    private fun convertToHSSFRichTextString(src: XSSFRichTextString): HSSFRichTextString {
        val hssfRichTextString = HSSFRichTextString(src.string)

        logger.info("convert xss to hss")
        logger.info("src text: {}", src.string)

        for (formatIndex in 0 until src.numFormattingRuns()) {
            val srcFont = src.getFontOfFormattingRun(formatIndex) ?: continue
            val formatStart = src.getIndexOfFormattingRun(formatIndex)
            val formatLength = src.getLengthOfFormattingRun(formatIndex)

            logger.info("src font: {}", srcFont)

            val targetFont = srcFont.copy(this.src.workbook, this.targetWorkbook)

            logger.info("target font: {}", targetFont)

            hssfRichTextString.applyFont(formatStart, formatStart + formatLength, targetFont)
        }

        return hssfRichTextString
    }

    private fun convertToXSSFRichTextString(src: HSSFRichTextString): XSSFRichTextString {
        val xssfRichTextString = XSSFRichTextString(src.string)

        for (i in 0 until src.string.count()) {
            val font = src.getFontAtIndex(i)
            val targetFont = fontMap[font.toInt()] ?: continue
            xssfRichTextString.applyFont(i, i + 1, targetFont)
        }

        return xssfRichTextString
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
