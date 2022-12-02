package top.kikt.excel

import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.jupiter.api.Test
import java.io.File

class XSSFWorkbookTest : ILogger {

    @Test
    fun getXssfSharedStringTable() {
        val wb = File("sample/src1.xlsx").toWorkbook() as XSSFWorkbook
        val sharedStringTable = wb.sharedStringSource
        for (sharedStringItem in sharedStringTable.sharedStringItems) {
            logger.debug("sharedStringItem: {}", sharedStringItem)
        }

        run {
            // sheet get font
            val sheet = wb.getSheetAt(0)
            val cell = sheet.getRow(12).getCellOrNull("G") as? XSSFCell ?: return@run
//            val xssfCellStyle = cell.cellStyle
            val value = cell.richStringCellValue

            logger.info("value: {}, value.hasFormatting(): {}", value, value.hasFormatting())

            logger.info("numFormattingRuns: {}", value.numFormattingRuns())

            for (formatIndex in 0 until value.numFormattingRuns()) {
                val formatStart = value.getIndexOfFormattingRun(formatIndex)
                val length = value.getLengthOfFormattingRun(formatIndex)

                val fontIndex = value.getIndexOfFormattingRun(formatIndex)
                wb.getFontAt(fontIndex).let {
                    logger.info("formatStart: {}, length: {}", formatStart, length)
                    logger.info("font: {}", it)
                }
            }
        }

        // styles
        run {
            val styles = wb.stylesSource
            for (font in styles.fonts) {
                logger.debug("font: {}", font)
            }
        }

    }

}