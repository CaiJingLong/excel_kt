@file:Suppress("unused")

package top.kikt.excel.writer

import org.junit.jupiter.api.Test
import top.kikt.excel.createParent
import java.io.File

internal class MergeExcelWriterTest {

    class Data1(
        @ExcelProperty("名字", 0)
        val name: String,

        @ExcelProperty("描述", 1)
        val desc: String,
    )

    class Data2(@ExcelProperty("密码", 0) val password: String)

    class Data3(@ExcelProperty("邮箱", 0) val email: String)

    class DataWrapper(
        @MergeExcelProperty(0, "数据1")
        val data1: Data1,

        @MergeExcelProperty(1, "数据2")
        val data2: Data2,

        @MergeExcelProperty(2, "数据3")
        val data3: Data3,
    )

    @Test
    fun write() {
        val data = DataWrapper(
            Data1("小明", "数学好"),
            Data2("123456"),
            Data3("xiaomi@example.com")
        )

        val data2 = DataWrapper(
            Data1("小红", "语文好"),
            Data2("456789"),
            Data3("xiaohong@example.com")
        )

        val file = File("sample/merge/merge_writer.xlsx").createParent()

        val writer = MergeExcelWriter(listOf(data, data2))
        file.outputStream().use {
            writer.write(it)
        }
    }

}