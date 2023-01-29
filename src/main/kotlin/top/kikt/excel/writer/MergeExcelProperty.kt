package top.kikt.excel.writer

@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
annotation class MergeExcelProperty(
    val value: String = "",
    val index: Int = Int.MAX_VALUE,
)