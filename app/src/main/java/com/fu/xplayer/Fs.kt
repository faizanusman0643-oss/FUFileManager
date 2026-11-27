package com.fu.xplayer

import android.content.ActivityNotFoundException
import android.content.Context
import android.content.Intent
import android.media.MediaScannerConnection
import android.net.Uri
import android.os.Environment
import android.os.StatFs
import android.provider.MediaStore
import android.webkit.MimeTypeMap
import android.widget.Toast
import androidx.core.content.FileProvider
import java.io.BufferedInputStream
import java.io.BufferedOutputStream
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipInputStream
import java.util.zip.ZipOutputStream

enum class Kind { FOLDER, VIDEO, AUDIO, IMAGE, DOC, ARCHIVE, APK, OTHER }

data class BinItem(val file: File, val original: String, val time: Long)

data class Drive(val label: String, val root: File, val total: Long, val free: Long)

object Fs {
    val root: File = Environment.getExternalStorageDirectory()
    val bin = File(root, ".FURecycle")
    private val binIndex get() = File(bin, ".index")
    val downloads: File = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)

    val VIDEO = setOf("mp4", "mkv", "avi", "mov", "wmv", "flv", "3gp", "webm", "m4v", "ts", "mpg", "mpeg", "m3u8", "vob", "rmvb", "f4v", "mts", "m2ts", "asf", "divx")
    val AUDIO = setOf("mp3", "wav", "aac", "m4a", "flac", "ogg", "opus", "wma", "amr", "mid", "midi", "3ga", "mka", "ac3", "m4b", "aiff")
    val IMAGE = setOf("jpg", "jpeg", "png", "gif", "bmp", "webp", "heic", "heif")
    val DOC = setOf("pdf", "doc", "docx", "xls", "xlsx", "xlsm", "ppt", "pptx", "txt", "rtf", "csv", "odt", "ods", "html", "htm", "xml", "json", "md", "log", "epub")
    val ARCH = setOf("zip", "rar", "7z", "tar", "gz", "bz2", "xz", "tgz")

    fun kind(f: File): Kind {
        if (f.isDirectory) return Kind.FOLDER
        return kindOfExt(f.extension.lowercase())
    }

    fun kindOfExt(e: String): Kind = when (e) {
        in VIDEO -> Kind.VIDEO
        in AUDIO -> Kind.AUDIO
        in IMAGE -> Kind.IMAGE
        in DOC -> Kind.DOC
        in ARCH -> Kind.ARCHIVE
        "apk", "xapk" -> Kind.APK
        else -> Kind.OTHER
    }

    fun isMedia(f: File) = kind(f).let { it == Kind.VIDEO || it == Kind.AUDIO }

    fun typeName(f: File): String = when {
        f.isDirectory -> "File folder"
        f.extension.isEmpty() -> "File"
        else -> f.extension.uppercase() + " " + when (kind(f)) {
            Kind.VIDEO -> "Video"; Kind.AUDIO -> "Audio"; Kind.IMAGE -> "Image"; Kind.DOC -> "Document"
            Kind.ARCHIVE -> "Archive"; Kind.APK -> "Application"; else -> "File"
        }
    }

    fun mime(f: File): String {
        val e = f.extension.lowercase()
        return when (e) {
            "m3u8" -> "application/x-mpegURL"
            "mkv" -> "video/x-matroska"
            "apk" -> "application/vnd.android.package-archive"
            else -> MimeTypeMap.getSingleton().getMimeTypeFromExtension(e) ?: "*/*"
        }
    }

    fun list(dir: File, hidden: Boolean): List<File> =
        (dir.listFiles() ?: emptyArray()).filter { hidden || !it.name.startsWith(".") }

    fun drives(ctx: Context): List<Drive> {
        val out = mutableListOf<Drive>()
        fun add(label: String, f: File) {
            try { val s = StatFs(f.path); out += Drive(label, f, s.totalBytes, s.availableBytes) } catch (_: Exception) {}
        }
        add("Local Disk (C:) — Internal storage", root)
        var letter = 'D'
        ctx.getExternalFilesDirs(null).drop(1).filterNotNull().forEach {
            val r = File(it.absolutePath.substringBefore("/Android/"))
            if (r.exists()) add("SD Card ($letter:)", r)
            letter++
        }
        return out
    }

    /** Windows style unique name: "name (2).ext" */
    fun unique(dir: File, name: String): File {
        var f = File(dir, name)
        if (!f.exists()) return f
        val dot = name.lastIndexOf('.')
        val base = if (dot > 0) name.substring(0, dot) else name
        val ext = if (dot > 0) name.substring(dot) else ""
        var i = 2
        while (f.exists()) { f = File(dir, "$base ($i)$ext"); i++ }
        return f
    }

    fun sizeOf(f: File): Long = if (f.isFile) f.length() else (f.listFiles() ?: emptyArray()).sumOf { sizeOf(it) }

    fun counts(f: File): Pair<Int, Int> {
        var files = 0; var dirs = 0
        f.walkTopDown().forEach { if (it != f) { if (it.isDirectory) dirs++ else files++ } }
        return files to dirs
    }

    fun countBytes(files: List<File>) = files.sumOf { sizeOf(it) }

    /** Copy (or move) files into dest; progress = bytes done / total. Returns created files. */
    fun transfer(files: List<File>, dest: File, cut: Boolean, onProg: (Long, Long, String) -> Unit): List<File> {
        val total = countBytes(files).coerceAtLeast(1)
        var done = 0L
        val out = mutableListOf<File>()
        val buf = ByteArray(256 * 1024)
        fun copyRec(src: File, dstDir: File, name: String = src.name): File {
            val target = unique(dstDir, name)
            if (src.isDirectory) {
                target.mkdirs()
                (src.listFiles() ?: emptyArray()).forEach { copyRec(it, target) }
            } else {
                onProg(done, total, src.name)
                BufferedInputStream(FileInputStream(src)).use { i ->
                    BufferedOutputStream(FileOutputStream(target)).use { o ->
                        while (true) {
                            val n = i.read(buf); if (n < 0) break
                            o.write(buf, 0, n); done += n
                            onProg(done, total, src.name)
                        }
                    }
                }
                target.setLastModified(src.lastModified())
            }
            return target
        }
        dest.mkdirs()
        for (f in files) {
            if (dest.absolutePath == f.absolutePath || dest.absolutePath.startsWith(f.absolutePath + "/")) continue
            if (cut) {
                val t = unique(dest, f.name)
                if (f.parentFile?.absolutePath == dest.absolutePath) { out += f; continue }
                if (f.renameTo(t)) { done += sizeOf(t); onProg(done, total, f.name); out += t; continue }
                val c = copyRec(f, dest); deleteRec(f); out += c
            } else out += copyRec(f, dest)
        }
        return out
    }

    fun deleteRec(f: File): Boolean = f.deleteRecursively()

    // ---------- Recycle Bin ----------
    fun toBin(files: List<File>): Int {
        bin.mkdirs()
        File(bin, ".nomedia").createNewFile()
        var n = 0
        val lines = StringBuilder()
        for (f in files) {
            val t = unique(bin, f.name)
            val ok = f.renameTo(t) || run {
                try { transfer(listOf(f), bin, false) { _, _, _ -> }; deleteRec(f); true } catch (_: Exception) { false }
            }
            if (ok) { n++; lines.append(t.name).append('\t').append(f.absolutePath).append('\t').append(System.currentTimeMillis()).append('\n') }
        }
        binIndex.appendText(lines.toString())
        return n
    }

    fun binItems(): List<BinItem> {
        val idx = HashMap<String, Pair<String, Long>>()
        if (binIndex.exists()) binIndex.readLines().forEach { l ->
            val p = l.split('\t'); if (p.size >= 3) idx[p[0]] = p[1] to (p[2].toLongOrNull() ?: 0L)
        }
        return (bin.listFiles() ?: emptyArray()).filter { !it.name.startsWith(".") }.map {
            val e = idx[it.name]
            BinItem(it, e?.first ?: File(root, it.name).absolutePath, e?.second ?: it.lastModified())
        }.sortedByDescending { it.time }
    }

    private fun rewriteIndex() {
        val keep = binItems()
        binIndex.writeText(keep.joinToString("") { "${it.file.name}\t${it.original}\t${it.time}\n" })
    }

    fun restore(items: List<BinItem>): Int {
        var n = 0
        for (b in items) {
            val o = File(b.original); o.parentFile?.mkdirs()
            val t = unique(o.parentFile ?: root, o.name)
            if (b.file.renameTo(t)) n++
        }
        rewriteIndex(); return n
    }

    fun purge(items: List<BinItem>) { items.forEach { deleteRec(it.file) }; rewriteIndex() }
    fun emptyBin() { binItems().forEach { deleteRec(it.file) }; binIndex.delete() }

    // ---------- Search ----------
    fun search(base: File, q: String, hidden: Boolean, limit: Int = 800): List<File> {
        val needle = q.trim().lowercase()
        if (needle.isEmpty()) return emptyList()
        val out = ArrayList<File>()
        base.walkTopDown().onEnter { hidden || !it.name.startsWith(".") || it == base }.forEach {
            if (out.size >= limit) return out
            if (it != base && it.name.lowercase().contains(needle) && (hidden || !it.name.startsWith("."))) out += it
        }
        return out
    }

    // ---------- Libraries (MediaStore) ----------
    fun library(ctx: Context, k: Kind): List<File> {
        val uri = MediaStore.Files.getContentUri("external")
        val col = MediaStore.Files.FileColumns.MEDIA_TYPE
        val sel = when (k) {
            Kind.VIDEO -> "$col=${MediaStore.Files.FileColumns.MEDIA_TYPE_VIDEO}"
            Kind.AUDIO -> "$col=${MediaStore.Files.FileColumns.MEDIA_TYPE_AUDIO}"
            Kind.IMAGE -> "$col=${MediaStore.Files.FileColumns.MEDIA_TYPE_IMAGE}"
            else -> "$col=${MediaStore.Files.FileColumns.MEDIA_TYPE_NONE}"
        }
        val out = ArrayList<File>()
        try {
            ctx.contentResolver.query(
                uri, arrayOf(MediaStore.MediaColumns.DATA), sel, null,
                MediaStore.MediaColumns.DATE_MODIFIED + " DESC"
            )?.use { c ->
                while (c.moveToNext()) {
                    val p = c.getString(0) ?: continue
                    val f = File(p)
                    if (k == Kind.DOC && kindOfExt(f.extension.lowercase()) != Kind.DOC) continue
                    if (f.isFile) out += f
                }
            }
        } catch (_: Exception) {}
        return out
    }

    fun scan(ctx: Context, paths: List<String>) {
        if (paths.isEmpty()) return
        try { MediaScannerConnection.scanFile(ctx, paths.toTypedArray(), null, null) } catch (_: Exception) {}
    }

    // ---------- Zip ----------
    fun zip(files: List<File>, out: File, onProg: (Long, Long, String) -> Unit) {
        val total = countBytes(files).coerceAtLeast(1); var done = 0L
        val buf = ByteArray(128 * 1024)
        ZipOutputStream(BufferedOutputStream(FileOutputStream(out))).use { z ->
            fun add(f: File, path: String) {
                if (f.isDirectory) {
                    z.putNextEntry(ZipEntry("$path/")); z.closeEntry()
                    (f.listFiles() ?: emptyArray()).forEach { add(it, "$path/${it.name}") }
                } else {
                    z.putNextEntry(ZipEntry(path))
                    FileInputStream(f).use { i -> while (true) { val n = i.read(buf); if (n < 0) break; z.write(buf, 0, n); done += n; onProg(done, total, f.name) } }
                    z.closeEntry()
                }
            }
            files.forEach { add(it, it.name) }
        }
    }

    fun unzip(zip: File, destParent: File, onProg: (Long, Long, String) -> Unit): File {
        val dest = unique(destParent, zip.nameWithoutExtension); dest.mkdirs()
        val total = zip.length().coerceAtLeast(1); var done = 0L
        val buf = ByteArray(128 * 1024)
        ZipInputStream(BufferedInputStream(FileInputStream(zip))).use { z ->
            while (true) {
                val e = z.nextEntry ?: break
                val t = File(dest, e.name)
                if (!t.canonicalPath.startsWith(dest.canonicalPath)) continue
                if (e.isDirectory) t.mkdirs() else {
                    t.parentFile?.mkdirs()
                    FileOutputStream(t).use { o -> while (true) { val n = z.read(buf); if (n < 0) break; o.write(buf, 0, n) } }
                    done += e.compressedSize.coerceAtLeast(0); onProg(done, total, e.name)
                }
            }
        }
        return dest
    }

    // ---------- Open / share ----------
    fun uriFor(ctx: Context, f: File): Uri = FileProvider.getUriForFile(ctx, ctx.packageName + ".fp", f)

    fun openExternal(ctx: Context, f: File, chooser: Boolean = false) {
        try {
            val i = Intent(Intent.ACTION_VIEW).setDataAndType(uriFor(ctx, f), mime(f))
                .addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_ACTIVITY_NEW_TASK)
            ctx.startActivity(if (chooser) Intent.createChooser(i, "Open with").addFlags(Intent.FLAG_ACTIVITY_NEW_TASK) else i)
        } catch (e: ActivityNotFoundException) {
            Toast.makeText(ctx, "Is file ko kholne wali koi app nahi mili", Toast.LENGTH_SHORT).show()
        } catch (e: Exception) {
            Toast.makeText(ctx, "Open fail: ${e.message}", Toast.LENGTH_SHORT).show()
        }
    }

    fun share(ctx: Context, files: List<File>) {
        val list = files.filter { it.isFile }
        if (list.isEmpty()) return
        val uris = ArrayList(list.map { uriFor(ctx, it) })
        val i = if (uris.size == 1) Intent(Intent.ACTION_SEND).putExtra(Intent.EXTRA_STREAM, uris[0]).setType(mime(list[0]))
        else Intent(Intent.ACTION_SEND_MULTIPLE).putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris).setType("*/*")
        i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        ctx.startActivity(Intent.createChooser(i, "Share").addFlags(Intent.FLAG_ACTIVITY_NEW_TASK))
    }
}

fun fmtSize(b: Long): String = when {
    b < 0 -> "?"
    b < 1024 -> "$b bytes"
    b < 1024L * 1024 -> "%.1f KB".format(b / 1024.0)
    b < 1024L * 1024 * 1024 -> "%.1f MB".format(b / 1048576.0)
    else -> "%.2f GB".format(b / 1073741824.0)
}
