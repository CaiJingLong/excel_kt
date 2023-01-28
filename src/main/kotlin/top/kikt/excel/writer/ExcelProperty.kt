package top.kikt.excel.writer

@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelProperty(
    val value: String,
    val index: Int
)
