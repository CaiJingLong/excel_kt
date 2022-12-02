package top.kikt.excel

interface ILogger {
    val logger: org.slf4j.Logger
        get() = org.slf4j.LoggerFactory.getLogger(this.javaClass)

}