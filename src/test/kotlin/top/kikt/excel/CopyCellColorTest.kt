package top.kikt.excel

import org.junit.jupiter.api.Test
import java.io.File

internal class CopyCellColorTest : ILogger {

    private val dir = File("sample/copy-cell/")
    private val src = File(dir, "2007.xlsx")

    @Test
    fun copyCellColor() {
        val dst07 = File(dir, "dst-07.xlsx")
        val dst03 = File(dir, "dst-03.xls")

        listOf(
            dst07,
            dst03,
        ).forEach { dstFile ->
            if (dstFile.exists()) {
                dstFile.delete()
            }

            val srcWb = src.toWorkbook()
            val dstWb = dstFile.createWorkbook(true)

            val sheet = srcWb.getSheetAt(0)
            //            sheet.copyTo(dstWb, active = true)
            val dstSheet = dstWb.createSheet()

            for (cellTag in arrayOf(
                "A1",
                "A4",
                "A5",
                "A6",
                "A7",
            )) {
                val cell = sheet.getCellOrCreate(cellTag)
                val dstCell = dstSheet.getCellOrCreate(cellTag)
                logger.info("src cell value: {}", cell.richStringCellValue.string)

                dstCell.setCellValue(cell.richStringCellValue.string)
                val cs = dstCell.workbook.createCellStyle()
                dstCell.cellStyle = cs
                cell.copyCellFont(dstCell)
            }

            dstWb.saveTo(dstFile)
        }
    }

}