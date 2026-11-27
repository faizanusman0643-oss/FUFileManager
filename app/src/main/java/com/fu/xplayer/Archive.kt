package com.fu.xplayer

import com.github.junrar.exception.UnsupportedRarV5Exception
import net.lingala.zip4j.exception.ZipException
import net.lingala.zip4j.model.ZipParameters
import net.lingala.zip4j.model.enums.AesKeyStrength
import net.lingala.zip4j.model.enums.EncryptionMethod
import org.apache.commons.compress.PasswordRequiredException
import org.apache.commons.compress.archivers.sevenz.SevenZFile
import org.apache.commons.compress.archivers.tar.TarArchiveInputStream
import org.apache.commons.compress.compressors.bzip2.BZip2CompressorInputStream
import org.apache.commons.compress.compressors.gzip.GzipCompressorInputStream
import org.apache.commons.compress.compressors.xz.XZCompressorInputStream
import java.io.BufferedInputStream
import java.io.File
import java.io.FileInputStream
import java.io.IOException
import java.io.InputStream

data class AEntry(val name: String, val size: Long, val dir: Boolean, val time: Long)

/** Thrown when the archive is locked and the password is missing or wrong. */
class NeedPassword(val wrong: Boolean) : IOException(if (wrong) "Password galat hai" else "Is archive pe password laga hai")

/** Read & extract ZIP, RAR (v4), 7Z, TAR, TAR.GZ/TGZ, TAR.BZ2, TAR.XZ, GZ — with password support. */
object Archive {
    fun type(f: File): String {
        val n = f.name.lowercase()
        return when {
            n.endsWith(".zip") || n.endsWith(".jar") || n.endsWith(".xapk") -> "zip"
            n.endsWith(".rar") -> "rar"
            n.endsWith(".7z") -> "7z"
            n.endsWith(".tar") -> "tar"
            n.endsWith(".tar.gz") || n.endsWith(".tgz") -> "tgz"
            n.endsWith(".tar.bz2") || n.endsWith(".tbz2") -> "tbz"
            n.endsWith(".tar.xz") || n.endsWith(".txz") -> "txz"
            n.endsWith(".gz") -> "gz"
            n.endsWith(".bz2") -> "bz2"
            n.endsWith(".xz") -> "xz"
            else -> ""
        }
    }

    fun supported(f: File) = type(f).isNotEmpty()

    private fun tarStream(f: File, t: String): TarArchiveInputStream {
        val raw: InputStream = BufferedInputStream(FileInputStream(f))
        return TarArchiveInputStream(
            when (t) {
                "tgz" -> GzipCompressorInputStream(raw)
                "tbz" -> BZip2CompressorInputStream(raw)
                "txz" -> XZCompressorInputStream(raw)
                else -> raw
            }
        )
    }

    private fun single(f: File, t: String): InputStream {
        val raw = BufferedInputStream(FileInputStream(f))
        return when (t) { "gz" -> GzipCompressorInputStream(raw); "bz2" -> BZip2CompressorInputStream(raw); else -> XZCompressorInputStream(raw) }
    }

    private fun zip(f: File, pw: String?) = net.lingala.zip4j.ZipFile(f, pw?.takeIf { it.isNotEmpty() }?.toCharArray())

    private fun seven(f: File, pw: String?): SevenZFile =
        if (pw.isNullOrEmpty()) SevenZFile(f) else SevenZFile(f, pw.toCharArray())

    private fun rar(f: File, pw: String?) =
        if (pw.isNullOrEmpty()) com.github.junrar.Archive(f) else com.github.junrar.Archive(f, pw)

    /** Convert library errors into friendly messages / NeedPassword. */
    private inline fun <T> guard(pw: String?, block: () -> T): T {
        try { return block() } catch (e: NeedPassword) { throw e } catch (e: Throwable) {
            val m = (e.message ?: "") + " " + (e.cause?.message ?: "")
            val had = !pw.isNullOrEmpty()
            when {
                e is UnsupportedRarV5Exception -> throw IOException("Ye RAR5 format hai — abhi sirf RAR4 support hai. PC pe WinRAR se ZIP/RAR4 bana lein.")
                e is PasswordRequiredException -> throw NeedPassword(had)
                e is ZipException && e.type == ZipException.Type.WRONG_PASSWORD -> throw NeedPassword(had)
                m.contains("password", true) || m.contains("encrypt", true) -> throw NeedPassword(had)
                had && m.contains("crc", true) -> throw NeedPassword(true)
                e is IOException -> throw e
                else -> throw IOException(e.message ?: e.javaClass.simpleName)
            }
        }
    }

    /** True if opening needs a password. */
    fun locked(f: File): Boolean = try {
        when (type(f)) {
            "zip" -> zip(f, null).isEncrypted
            "rar" -> rar(f, null).use { a -> a.isEncrypted || a.fileHeaders.any { it.isEncrypted } }
            "7z" -> { seven(f, null).use { z -> z.entries.count() }; false }
            else -> false
        }
    } catch (e: PasswordRequiredException) { true } catch (e: Throwable) {
        (e.message ?: "").contains("password", true)
    }

    fun list(f: File, pw: String? = null): List<AEntry> = guard(pw) {
        when (val t = type(f)) {
            "zip" -> zip(f, pw).fileHeaders.map { AEntry(it.fileName, it.uncompressedSize, it.isDirectory, it.lastModifiedTimeEpoch) }
            "7z" -> seven(f, pw).use { z -> z.entries.map { AEntry(it.name, it.size, it.isDirectory, if (it.hasLastModifiedDate) it.lastModifiedDate.time else 0L) } }
            "tar", "tgz", "tbz", "txz" -> tarStream(f, t).use { s ->
                val out = ArrayList<AEntry>()
                while (true) { val e = s.nextEntry ?: break; out += AEntry(e.name, e.size, e.isDirectory, e.lastModifiedDate?.time ?: 0L) }
                out
            }
            "gz", "bz2", "xz" -> listOf(AEntry(f.name.substringBeforeLast('.'), -1, false, f.lastModified()))
            "rar" -> rar(f, pw).use { a ->
                if (a.isEncrypted && pw.isNullOrEmpty()) throw NeedPassword(false)
                a.fileHeaders.map { AEntry(it.fileName.replace('\\', '/'), it.fullUnpackSize, it.isDirectory, 0L) }
            }
            else -> throw IOException("Ye archive format support nahi")
        }
    }

    private fun safe(dest: File, name: String): File? {
        val t = File(dest, name)
        return if (t.canonicalPath.startsWith(dest.canonicalPath)) t else null
    }

    private fun write(i: InputStream, t: File, onProg: (String) -> Unit, name: String) {
        t.parentFile?.mkdirs(); onProg(name)
        t.outputStream().buffered().use { i.copyTo(it, 128 * 1024) }
    }

    /** Extract everything (only = null) or just entries named in [only]. Returns output folder. */
    fun extract(f: File, destParent: File, only: Set<String>? = null, pw: String? = null, onProg: (String) -> Unit = {}): File = guard(pw) {
        val t = type(f)
        val base = f.name.removeSuffix(".gz").removeSuffix(".bz2").removeSuffix(".xz").substringBeforeLast('.')
        val dest = if (only != null) destParent else Fs.unique(destParent, base)
        dest.mkdirs()
        fun want(n: String) = only == null || n in only
        fun target(o: File) = if (only != null) Fs.unique(dest, o.name) else o
        try {
            when (t) {
                "zip" -> {
                    val z = zip(f, pw)
                    if (z.isEncrypted && pw.isNullOrEmpty()) throw NeedPassword(false)
                    for (h in z.fileHeaders) {
                        val o = safe(dest, h.fileName) ?: continue
                        if (!want(h.fileName)) continue
                        if (h.isDirectory) o.mkdirs() else z.getInputStream(h).use { write(it, target(o), onProg, h.fileName) }
                    }
                }
                "7z" -> seven(f, pw).use { z ->
                    while (true) {
                        val e = z.nextEntry ?: break
                        val o = safe(dest, e.name) ?: continue
                        if (!want(e.name)) continue
                        if (e.isDirectory) o.mkdirs() else z.getInputStream(e).use { write(it, target(o), onProg, e.name) }
                    }
                }
                "tar", "tgz", "tbz", "txz" -> tarStream(f, t).use { s ->
                    while (true) {
                        val e = s.nextEntry ?: break
                        val o = safe(dest, e.name) ?: continue
                        if (!want(e.name)) continue
                        if (e.isDirectory) o.mkdirs() else write(s, target(o), onProg, e.name)
                    }
                }
                "gz", "bz2", "xz" -> single(f, t).use { write(it, Fs.unique(dest, f.name.substringBeforeLast('.')), onProg, f.name) }
                "rar" -> rar(f, pw).use { a ->
                    if ((a.isEncrypted || a.fileHeaders.any { it.isEncrypted }) && pw.isNullOrEmpty()) throw NeedPassword(false)
                    for (h in a.fileHeaders) {
                        val n = h.fileName.replace('\\', '/')
                        val o = safe(dest, n) ?: continue
                        if (!want(n)) continue
                        if (h.isDirectory) { o.mkdirs(); continue }
                        val out = target(o); out.parentFile?.mkdirs(); onProg(n)
                        out.outputStream().buffered().use { a.extractFile(h, it) }
                    }
                }
                else -> throw IOException("Ye archive format support nahi")
            }
        } catch (e: Throwable) {
            if (only == null && (dest.listFiles()?.isEmpty() != false)) dest.deleteRecursively()
            throw e
        }
        dest
    }

    /** Extract one entry to cache so it can be opened/played directly. */
    fun extractToCache(f: File, entry: String, cache: File, pw: String? = null): File {
        val dir = File(cache, "arc_open").also { it.deleteRecursively(); it.mkdirs() }
        extract(f, dir, setOf(entry), pw)
        return dir.listFiles()?.firstOrNull() ?: throw IOException("Extract fail")
    }

    /** Create a ZIP (optionally AES-256 password protected). */
    fun createZip(files: List<File>, out: File, pw: String?, onProg: (String) -> Unit) {
        val z = net.lingala.zip4j.ZipFile(out, pw?.takeIf { it.isNotEmpty() }?.toCharArray())
        val p = ZipParameters().apply {
            if (!pw.isNullOrEmpty()) {
                isEncryptFiles = true
                encryptionMethod = EncryptionMethod.AES
                aesKeyStrength = AesKeyStrength.KEY_STRENGTH_256
            }
        }
        for (f in files) { onProg(f.name); if (f.isDirectory) z.addFolder(f, p) else z.addFile(f, p) }
        z.close()
    }
}
