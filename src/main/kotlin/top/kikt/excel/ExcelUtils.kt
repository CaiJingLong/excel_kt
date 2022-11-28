package top.kikt.excel

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.IOException

@Suppress("unused")
object ExcelUtils {

    @JvmStatic
    fun getWorkbook(path: String): Workbook = getWorkbook(File(path))

    @JvmStatic
    fun getWorkbook(file: File): Workbook {
        if (!file.exists()) {
            throw IOException("文件 ${file.path} 不存在")
        }

        return try {
            HSSFWorkbook(file.inputStream())
        } catch (e: Exception) {
            try {
                XSSFWorkbook(file.inputStream())
            } catch (e: Exception) {
                e.printStackTrace()
                throw IOException("The file ${file.absolutePath} 创建失败")
            }
        }
    }

}