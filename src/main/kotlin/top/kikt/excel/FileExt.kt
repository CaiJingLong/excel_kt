package top.kikt.excel

import java.io.File

/**
 * String to file.
 */
fun String.toFile(): File {
    return File(this)
}

/**
 * If file is not exist, create parent dir and it.
 */
fun File.createIfNotExists(): File {
    if (!exists()) {
        parentFile.mkdirs()
        createNewFile()
    }
    return this
}

/**
 * If file is not exist, create parent dir and it.
 */
fun String.createIfNotExists(): String {
    toFile().createIfNotExists()
    return this
}

/**
 * Create parent dir.
 */
fun File.createParent(): File {
    parentFile.mkdirs()
    return this
}