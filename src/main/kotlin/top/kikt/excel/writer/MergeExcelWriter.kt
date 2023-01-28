package top.kikt.excel.writer

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.OutputStream
import java.lang.reflect.Field

class MergeExcelWriter(val data: List<Any>) {

    private class MergeProperty {
        var wrapperIndex: Int = -1
        var propertyIndex: Int = -1
        lateinit var wrapperField: Field
        lateinit var field: Field
        lateinit var mergeExcelProperty: MergeExcelProperty

        fun getValue(data: Any): Any? {
            wrapperField.isAccessible = true
            val wrapperValue = wrapperField.get(data)
            field.isAccessible = true
            return field.get(wrapperValue)
        }
    }

    fun write(outputStream: OutputStream) {
        val list = mutableListOf<MergeProperty>()

        if (data.isEmpty()) {
            return
        }

        val first = data[0]

        // handle first data to get title and merge type
        val clazz = first.javaClass
        clazz.declaredFields.forEach { wrapperField ->
            val annotation = wrapperField.getAnnotation(MergeExcelProperty::class.java)
            if (annotation != null) {
                val mergeIndex = annotation.index
                wrapperField.isAccessible = true

                for (field in wrapperField.type.declaredFields) {
                    val mergeProperty = MergeProperty()

                    val excelProperty = field.getAnnotation(ExcelProperty::class.java) ?: continue

                    val propertyIndex = excelProperty.index

                    mergeProperty.wrapperIndex = mergeIndex
                    mergeProperty.propertyIndex = propertyIndex
                    mergeProperty.wrapperField = wrapperField
                    mergeProperty.field = field
                    mergeProperty.mergeExcelProperty = annotation
                    list.add(mergeProperty)
                }
            }
        }

        list.apply {
            sortBy { it.propertyIndex }
            sortBy { it.wrapperIndex }
        }

        val workbook = WorkbookFactory.create(true)
        val sheet = workbook.createSheet("sheet1")

        writeTitle(list, sheet)
        writeData(list, sheet)

        workbook.write(outputStream)
    }

    private fun writeData(propertyList: MutableList<MergeProperty>, sheet: Sheet) {
        for (i in 1..data.size) {
            val row = sheet.createRow(sheet.lastRowNum + 1)
            val item = data[i - 1]
            for (mergeProperty in propertyList) {
                val value = mergeProperty.getValue(item)
                val lastCellNum = row.lastCellNum.toInt()
                val cell =
                    if (lastCellNum == -1) {
                        row.createCell(0)
                    } else {
                        row.createCell(lastCellNum)
                    }

                val typeAnno = mergeProperty.field.getAnnotation(ExcelTypeAnno::class.java)
                if (typeAnno != null) {
                    when (typeAnno.value) {
                        ExcelType.DATE -> cell.setCellValue(value.toString())
                        ExcelType.STRING -> cell.setCellValue(value.toString())
                        ExcelType.NUMBER -> cell.setCellValue(value.toString().toDoubleOrNull() ?: 0.0)
                        ExcelType.BOOLEAN -> cell.setCellValue(value.toString().toBoolean())
                        ExcelType.FORMULA -> cell.setCellValue(value.toString())
                        ExcelType.BLANK -> cell.setCellValue(value.toString())
                        ExcelType.ERROR -> cell.setCellValue(value.toString())
                    }
                    continue
                }

                when (value) {
                    is Int -> cell.setCellValue(value.toDouble())
                    is Double -> cell.setCellValue(value)
                    is String -> cell.setCellValue(value)
                    else -> cell.setCellValue(value.toString())
                }
            }
        }
    }

    private fun writeTitle(list: MutableList<MergeProperty>, sheet: Sheet) {
        run {
            val title = sheet.createRow(0)
            val wrapperMap = list.groupBy { it.wrapperIndex }
            val keys = wrapperMap.keys.sortedBy { it }
            for ((index, entry) in keys.withIndex()) {
                val currentListCount = wrapperMap[entry]!!.size
                val start = if (index == 0) {
                    0
                } else {
                    var count = 0
                    for (i in 0 until index) {
                        count += wrapperMap[keys[i]]!!.size
                    }
                    count
                }
                val end = start + currentListCount

                val cell = title.createCell(start)
                val mergeProperty = wrapperMap[entry]!![0]
                val value = mergeProperty.mergeExcelProperty.name.ifEmpty { mergeProperty.wrapperField.name }

                cell.setCellValue(value)

                // merge cell
                if (currentListCount > 1) {
                    sheet.addMergedRegion(org.apache.poi.ss.util.CellRangeAddress(0, 0, start, end - 1))
                }
            }
        }

        val subTitle = sheet.createRow(1)
        for (mergeProperty in list) {
            mergeProperty.field.isAccessible = true
            var name = mergeProperty.field.getAnnotation(ExcelProperty::class.java).value
            if (name.isEmpty()) {
                name = mergeProperty.field.name
            }
            val lastCellNum = subTitle.lastCellNum.toInt()
            if (lastCellNum == -1) {
                subTitle.createCell(0).setCellValue(name)
                continue
            }
            subTitle.createCell(lastCellNum).setCellValue(name)
        }
    }

}
