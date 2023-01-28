package top.kikt.excel.writer

@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
annotation class MergeExcelProperty(val index: Int, val name: String = "")