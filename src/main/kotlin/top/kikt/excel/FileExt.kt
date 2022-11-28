package top.kikt.excel

import java.io.File

fun String.toFile(): File {
    return File(this)
}

/**
 * 如果文件不存在则创建文件和父目录
 */
fun File.createIfExists(): File {
    if (!exists()) {
        parentFile.mkdirs()
        createNewFile()
    }
    return this
}

fun String.createFileIfExists(): String {
    toFile().createIfExists()
    return this
}