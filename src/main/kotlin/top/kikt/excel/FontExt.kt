package top.kikt.excel

import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook

fun Font.copy(srcWorkbook: Workbook, targetWorkbook: Workbook): Font {
    val result = targetWorkbook.createFont().also {
        it.fontName = this.fontName
        it.fontHeight = this.fontHeight
        it.fontHeightInPoints = this.fontHeightInPoints
        it.italic = this.italic
        it.strikeout = this.strikeout
        it.typeOffset = this.typeOffset
        it.underline = this.underline
        it.charSet = this.charSet
        it.bold = this.bold
    }

    val rgbList = ArrayList<Int>() // r,g,b

    if (this is XSSFFont) {
        val argbHex = this.xssfColor.argbHex // "FFFF0000"
        val r = argbHex.substring(2, 4).toInt(16)
        val g = argbHex.substring(4, 6).toInt(16)
        val b = argbHex.substring(6, 8).toInt(16)
        rgbList.add(r)
        rgbList.add(g)
        rgbList.add(b)

    } else if (this is HSSFFont) {
        val hssfColor = this.getHSSFColor(srcWorkbook as HSSFWorkbook)
        val rgb = hssfColor.triplet // short array, If the color is red, return 255,0,0
        val r = rgb[0].toInt()
        val g = rgb[1].toInt()
        val b = rgb[2].toInt()
        rgbList.add(r)
        rgbList.add(g)
        rgbList.add(b)
    }

    if (rgbList.count() != 3) {
        return result
    }

    // Set result color
    if (targetWorkbook is HSSFWorkbook) {
        val customPalette = targetWorkbook.customPalette
        val hssfColor = customPalette.findColor(rgbList[0].toByte(), rgbList[1].toByte(), rgbList[2].toByte())
        if (hssfColor != null) {
            result.color = hssfColor.index
        } else {
            val newColor = customPalette.addColor(rgbList[0].toByte(), rgbList[1].toByte(), rgbList[2].toByte())
            result.color = newColor.index
        }
    } else if (targetWorkbook is XSSFWorkbook) {
        val color = XSSFColor(rgbList.map { it.toByte() }.toByteArray())
        (result as XSSFFont).setColor(color)
    }

    return result
}