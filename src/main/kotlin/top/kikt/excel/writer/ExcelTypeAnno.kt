package top.kikt.excel.writer

@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelTypeAnno(
    val value: ExcelType
)