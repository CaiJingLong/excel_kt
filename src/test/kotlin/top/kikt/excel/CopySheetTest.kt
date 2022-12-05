package top.kikt.excel

import org.junit.jupiter.api.Test
import org.slf4j.LoggerFactory
import java.io.File

internal class CopySheetTest {

    private val logger = LoggerFactory.getLogger(this.javaClass)

    private val dir = File("sample/copy-sheet")

    @Test
    fun xlsxCopyTest() {
        val wb1 = File(dir, "src1.xlsx").toWorkbook()
        val wb2 = File(dir, "src2.xlsx").toWorkbook()

        val outputFile = File(dir, "output.xlsx").createIfNotExists()

        val srcSheet = wb1.getSheetAt(0)
        val targetSheet = srcSheet.copyTo(wb2, active = true)

        logger.debug("targetSheet: {}", targetSheet.sheetName)

        wb2.saveTo(outputFile)
    }

    @Test
    fun xlsxToXls() {
        val wb1 = File(dir, "src1.xlsx").toWorkbook()
        val wb2 = File(dir, "src3.xls").toWorkbook()
        val outputFile = File(dir, "output-97-2007.xlsx").createIfNotExists()

        wb2.getSheetAt(0).copyTo(wb1, index = 0, name = "copied_sheet", active = true)

        wb1.saveTo(outputFile)
    }

    @Test
    fun xlsToXlsx() {
        val wb1 = File(dir, "src3.xls").toWorkbook()
        val wb2 = File(dir, "src1.xlsx").toWorkbook()
        val outputFile = File(dir, "output-07-to-97.xls").createIfNotExists()

        wb2.getSheetAt(0).copyTo(wb1, index = 0, name = "copied_sheet", active = true)

        wb1.saveTo(outputFile)
    }

}