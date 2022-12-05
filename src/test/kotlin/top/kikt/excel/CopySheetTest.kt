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
    fun xlsCopyTest() {
        val wb1 = File(dir, "src3.xls").toWorkbook()
        val wb2 = File(dir, "src4.xls").toWorkbook()

        val outputFile = File(dir, "xls-to-xls-output.xls").createIfNotExists()

        val srcSheet = wb1.getSheetAt(0)
        val targetSheet = srcSheet.copyTo(wb2, active = true)

        logger.info("copy xls to xls")
        logger.info("targetSheet: {}", targetSheet.sheetName)

        wb2.saveTo(outputFile)
    }

    @Test
    fun xlsxToXls() {
        val wb1 = File(dir, "src1.xlsx").toWorkbook()
        val wb2 = File(dir, "src3.xls").toWorkbook()
        val outputFile = File(dir, "xls-to-xlsx.xlsx").createIfNotExists()

        wb2.getSheetAt(0).copyTo(wb1, index = 0, name = "copied_sheet", active = true)

        wb1.saveTo(outputFile)
    }

    @Test
    fun xlsToXlsx() {
        val wb1 = File(dir, "src3.xls").toWorkbook()
        val wb2 = File(dir, "src1.xlsx").toWorkbook()
        val outputFile = File(dir, "xlsx-to-xls.xls").createIfNotExists()

        wb2.getSheetAt(0).copyTo(wb1, index = 0, name = "copied_sheet", active = true)

        wb1.saveTo(outputFile)
    }

}