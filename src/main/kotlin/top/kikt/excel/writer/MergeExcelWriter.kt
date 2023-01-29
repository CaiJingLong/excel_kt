package top.kikt.excel.writer

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.CellRangeAddress
import java.io.OutputStream
import java.lang.reflect.Field

class MergeExcelWriter(val data: List<Any>) {

    private class MergeProperty {
        var wrapperIndex: Int = -1
        var propertyIndex: Int = -1
        lateinit var wrapperField: Field
        lateinit var field: Field
        lateinit var mergeExcelProperty: MergeExcelProperty
        var dataTypeIndex = -1

        var columnIndex = -1
        var titleIndex = -1

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

        var dataTypeIndex = 0
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
                    mergeProperty.dataTypeIndex = dataTypeIndex

                    list.add(mergeProperty)
                }

                dataTypeIndex++
            }
        }

        list.apply {
            sortBy { it.propertyIndex }
            sortBy { it.dataTypeIndex }
            sortBy { it.wrapperIndex }
        }

        for ((index, mergeProperty) in list.withIndex()) {
            mergeProperty.columnIndex = index
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
                val cell = if (lastCellNum == -1) {
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
            val row = sheet.createRow(0)
            val groupData = list.groupBy { it.dataTypeIndex }

            groupData.values.forEach { mergePropertyList ->
                val min = mergePropertyList.minOf { it.columnIndex }
                val max = mergePropertyList.maxOf { it.columnIndex }
                val cell = row.createCell(min)
                cell.setCellValue(mergePropertyList[0].mergeExcelProperty.value)
                if (min != max) {
                    sheet.addMergedRegion(CellRangeAddress(0, 0, min, max))
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
            subTitle.createCell(mergeProperty.columnIndex).setCellValue(name)
        }
    }

}
