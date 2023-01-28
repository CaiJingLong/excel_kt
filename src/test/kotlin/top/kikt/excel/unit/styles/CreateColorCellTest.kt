package top.kikt.excel.unit.styles

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.junit.jupiter.api.Test
import top.kikt.excel.createWorkbook
import top.kikt.excel.getCellOrCreate
import top.kikt.excel.saveTo
import java.io.File

class CreateColorCellTest {

    @Test
    fun createColorCell07() {
        val file = File("sample/styles/color-cell-07.xlsx")
        if (file.exists()) {
            file.delete()
        }
        val wb = file.createWorkbook()
        val sheet = wb.createSheet()
        run {
            val cell = sheet.getCellOrCreate("A1")
            cell.setCellValue("Hello")
            cell.cellStyle = wb.createCellStyle()

            val font = wb.createFont()
            font.color = 0x0c

            cell.cellStyle.setFont(font)
        }

        run {
            val cell = sheet.getCellOrCreate("A2")
            cell.setCellValue("Hello")
            cell.cellStyle = wb.createCellStyle()
            val font = wb.createFont()
            // Set yellow
            font.color = HSSFColor.HSSFColorPredefined.SEA_GREEN.index
            cell.cellStyle.setFont(font)
        }

        run {
            // color rgb = 0x959697
            val cell = sheet.getCellOrCreate("A3")
            cell.setCellValue("Hello")
            cell.cellStyle = wb.createCellStyle()
            val font = wb.createFont() as XSSFFont
            // Set color rgb = 0x959697
//            wb as XSSFWorkbook
            val color = XSSFColor(byteArrayOf(0x95.toByte(), 0x96.toByte(), 0x97.toByte()))
            font.setColor(color)
            cell.cellStyle.setFont(font)
        }

        wb.saveTo(file)
    }

    @Test
    fun createColorCell03() {
        val file = File("sample/styles/color-cell-03.xls")
        if (file.exists()) {
            file.delete()
        }
        val wb = file.createWorkbook() as HSSFWorkbook
        val sheet = wb.createSheet()

        run {
            val cell = sheet.getCellOrCreate("A1")
            cell.setCellValue("Hello")
            cell.cellStyle = wb.createCellStyle()
            val font = wb.createFont()
            font.color = 0x0c
            cell.cellStyle.setFont(font)
        }

        run {
            val cell = sheet.getCellOrCreate("A2")
            cell.setCellValue("Hello")
            cell.cellStyle = wb.createCellStyle()
            val font = wb.createFont()
            // Set yellow
            font.color = HSSFColor.HSSFColorPredefined.SEA_GREEN.getIndex()
            cell.cellStyle.setFont(font)
        }

        run {
            // color rgb = 0x959697
            val cell = sheet.getCellOrCreate("A3")
            cell.setCellValue("Hello")
            cell.cellStyle = wb.createCellStyle()
            val font = wb.createFont()
            // Set color rgb = 0x959697
            val color = wb.customPalette.findSimilarColor(0x95.toByte(), 0x96.toByte(), 0x97.toByte())
            font.color = color.index
            cell.cellStyle.setFont(font)
        }

        wb.saveTo(file)
    }

}