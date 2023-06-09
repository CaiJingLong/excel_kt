package top.kikt.excel.unit.styles

import org.junit.jupiter.api.Test
import top.kikt.excel.*
import java.io.File
import kotlin.test.assertEquals

class FillBorderTest {

    @Test
    fun getIndexRangeTest() {
        val wb = File("sample/fill-border/fill-border.xlsx").toWorkbook()
        val sheet = wb.getSheetAt(2)

        assertEquals(sheet.getFirstNotNullRowIndex(), 1)
        assertEquals(sheet.getLastNotNullRowIndex(), 12)

        assertEquals(sheet.getFirstNotNullColumnIndex(), 0)
        assertEquals(sheet.getLastNotNullColumnIndex(), letterToColumnIndex("I"))
    }


    @Test
    fun sheetFillBorderTest() {
        val wb = File("sample/fill-border/fill-border.xlsx").toWorkbook()

        val sheet1 = wb.getSheetAt(0)

        sheet1.fillBorder("A1".toAddress(), "D4".toAddress())

        wb.saveTo(File("sample/fill-border/fill-border-result.xlsx"))
    }
}