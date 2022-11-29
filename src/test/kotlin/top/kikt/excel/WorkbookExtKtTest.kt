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

        wb2.saveTo(outputFile)
    }

    @Test
    fun sheet97To2007() {
        val wb1 = File("sample/src1.xlsx").toWorkbook()
        val wb2 = File("sample/src3.xls").toWorkbook()
        val outputFile = File("sample/output-97-2007.xlsx").createIfNotExists()

        wb2.getSheetAt(0).copyTo(wb1)

        wb1.saveTo(outputFile)
    }

}