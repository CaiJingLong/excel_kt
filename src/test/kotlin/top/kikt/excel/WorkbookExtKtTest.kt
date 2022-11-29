package top.kikt.excel

import org.junit.jupiter.api.Test
import org.slf4j.LoggerFactory
import java.io.File

internal class WorkbookExtKtTest {

    private val logger = LoggerFactory.getLogger(this.javaClass)

    @Test
    fun sheetCopyTo() {
        val wb1 = File("sample/src1.xlsx").toWorkbook()
        val wb2 = File("sample/src2.xlsx").toWorkbook()

        val outputFile = File("sample/output.xlsx").createIfNotExists()

        val srcSheet = wb1.getSheetAt(0)
        val targetSheet = srcSheet.copyTo(wb2)

        logger.debug("targetSheet: {}", targetSheet.sheetName)
        val firstCell = targetSheet.getRow(0).getCell(0)
        logger.debug("firstCell style: {}", firstCell.cellStyle.debugInfo())

        wb2.saveTo(outputFile)
    }

}