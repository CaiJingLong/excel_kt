package top.kikt.excel

import java.io.File

fun String.toFile(): File {
    return File(this)
}

/**
 * 如果文件不存在则创建文件和父目录
 */
fun File.createIfNotExists(): File {
    if (!exists()) {
        parentFile.mkdirs()
        createNewFile()
    }
    return this
}

fun String.createIfNotExists(): String {
    toFile().createIfNotExists()
    return this
}