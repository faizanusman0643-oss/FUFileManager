package com.fu.fm

import java.io.File
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

enum class Cat(val label: String) {
    PHOTO("Photos"), VIDEO("Videos"), AUDIO("Audio"), DOC("Docs"), OTHER("Other")
}

private val PHOTO_EXT = setOf("jpg", "jpeg", "png", "gif", "webp", "bmp", "heic", "heif")
private val VIDEO_EXT = setOf("mp4", "mkv", "avi", "mov", "3gp", "webm", "flv", "wmv", "m4v", "ts", "mpeg", "mpg")
private val AUDIO_EXT = setOf("mp3", "m4a", "aac", "wav", "flac", "ogg", "opus", "amr", "wma", "mid")
private val DOC_EXT = setOf("pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "txt", "csv", "rtf", "odt", "ods", "xlsm")

fun catOf(name: String): Cat {
    val e = name.substringAfterLast('.', "").lowercase()
    return when {
        e in PHOTO_EXT -> Cat.PHOTO
        e in VIDEO_EXT -> Cat.VIDEO
        e in AUDIO_EXT -> Cat.AUDIO
        e in DOC_EXT -> Cat.DOC
        else -> Cat.OTHER
    }
}

data class Item(val path: String, val name: String, val size: Long, val mtime: Long, val cat: Cat) {
    val file: File get() = File(path)
}

fun File.toItem() = Item(absolutePath, name, length(), lastModified(), catOf(name))

fun fmtSize(b: Long): String {
    if (b < 1024) return "$b B"
    val u = arrayOf("KB", "MB", "GB", "TB")
    var v = b.toDouble()
    var i = -1
    while (v >= 1024 && i < u.lastIndex) { v /= 1024; i++ }
    return String.format(Locale.US, "%.1f %s", v, u[i])
}

fun fmtDate(ms: Long): String = SimpleDateFormat("dd MMM yyyy", Locale.getDefault()).format(Date(ms))
