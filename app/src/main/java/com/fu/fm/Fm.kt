package com.fu.fm

import android.content.Context
import android.media.MediaScannerConnection
import android.os.Environment
import java.io.File
import java.security.MessageDigest
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

enum class SortMode { FORMAT, DATE, MONTH }

data class Moved(val from: File, val to: File)

object Fm {
    val root: File get() = Environment.getExternalStorageDirectory()

    /** .m3u8 / .mpd wale folder (HLS/DASH) ek unit hain: segments ke saath hi move hone chahiye. */
    private val BUNDLE_EXT = setOf("m3u8", "mpd")
    fun isBundle(d: File): Boolean =
        d.isDirectory && (d.listFiles()?.any { it.isFile && it.extension.lowercase() in BUNDLE_EXT } == true)

    /** Poori storage scan (hidden folders aur Android/ skip). */
    fun scan(): List<Item> {
        val out = ArrayList<Item>()
        fun walk(d: File, top: Boolean) {
            val list = d.listFiles() ?: return
            if (!top && list.any { it.isFile && it.extension.lowercase() in BUNDLE_EXT }) {
                list.filter { it.isFile && it.extension.lowercase() in BUNDLE_EXT }.forEach { out.add(it.toItem()) }
                return
            }
            for (f in list) {
                if (f.name.startsWith(".")) continue
                if (f.isDirectory) {
                    if (top && f.name == "Android") continue
                    walk(f, false)
                } else {
                    out.add(f.toItem())
                }
            }
        }
        walk(root, true)
        return out
    }

    fun unique(dir: File, name: String): File {
        var t = File(dir, name)
        if (!t.exists()) return t
        val base = if (name.contains('.')) name.substringBeforeLast('.') else name
        val ext = if (name.contains('.')) "." + name.substringAfterLast('.') else ""
        var i = 1
        while (t.exists()) {
            t = File(dir, "${base}_$i$ext")
            i++
        }
        return t
    }

    fun copyAny(src: File, dest: File): Boolean = try {
        dest.parentFile?.mkdirs()
        if (src.isDirectory) src.copyRecursively(dest, false) else { src.copyTo(dest, false); true }
    } catch (e: Exception) { false }

    fun moveFile(src: File, dest: File): Boolean {
        dest.parentFile?.mkdirs()
        if (src.renameTo(dest)) return true
        val ok = copyAny(src, dest)
        if (ok) src.deleteRecursively()
        return ok
    }

    fun moveAll(
        files: List<File>, destDir: File, copy: Boolean = false,
        onProgress: (Int, Int, String) -> Unit = { _, _, _ -> }
    ): List<Moved> {
        destDir.mkdirs()
        val res = ArrayList<Moved>()
        val units = files.map { f ->
            val p = f.parentFile
            if (f.isFile && f.extension.lowercase() in BUNDLE_EXT && p != null &&
                p.absolutePath != root.absolutePath && isBundle(p)) p else f
        }.distinct()
        for ((i, f) in units.withIndex()) {
            onProgress(i, units.size, f.name)
            if (!f.exists()) continue
            if (f.isDirectory && (destDir.absolutePath + "/").startsWith(f.absolutePath + "/")) continue
            val t = unique(destDir, f.name)
            val ok = if (copy) copyAny(f, t) else moveFile(f, t)
            if (ok) res.add(Moved(f, t))
        }
        onProgress(units.size, units.size, "")
        return res
    }

    fun deleteAll(
        files: List<File>,
        onProgress: (Int, Int, String) -> Unit = { _, _, _ -> }
    ): List<String> {
        val done = ArrayList<String>()
        for ((i, f) in files.withIndex()) {
            onProgress(i, files.size, f.name)
            if (f.deleteRecursively()) done.add(f.absolutePath)
        }
        onProgress(files.size, files.size, "")
        return done
    }

    fun rescan(ctx: Context, paths: Collection<String>) {
        if (paths.isNotEmpty()) MediaScannerConnection.scanFile(ctx, paths.toTypedArray(), null, null)
    }

    // ---------- Sort (bat: By Format / Date / Month, 8am cutoff) ----------
    fun sortFolder(
        dir: File, mode: SortMode,
        onProgress: (Int, Int, String) -> Unit = { _, _, _ -> }
    ): List<Moved> {
        val fmtD = SimpleDateFormat("yyyy-MM-dd", Locale.US)
        val fmtM = SimpleDateFormat("MM-MMMM-yyyy", Locale.ENGLISH)
        val shift = 8 * 3600_000L
        val res = ArrayList<Moved>()
        if (isBundle(dir)) return res
        val files = (dir.listFiles() ?: return res).filter { it.isFile && !it.name.startsWith(".") }
        for ((i, f) in files.withIndex()) {
            onProgress(i, files.size, f.name)
            val folder = when (mode) {
                SortMode.FORMAT -> {
                    val e = f.extension.uppercase()
                    if (e.isEmpty()) "No Extension" else "$e Files"
                }
                SortMode.DATE -> fmtD.format(Date(f.lastModified() - shift))
                SortMode.MONTH -> fmtM.format(Date(f.lastModified() - shift))
            }
            val t = unique(File(dir, folder), f.name)
            if (moveFile(f, t)) res.add(Moved(f, t))
        }
        onProgress(files.size, files.size, "")
        return res
    }

    // ---------- Consolidate (bat: Reverse Data) ----------
    fun consolidate(
        src: File, dest: File,
        onProgress: (Int, Int, String) -> Unit = { _, _, _ -> }
    ): Pair<List<Moved>, Int> {
        dest.mkdirs()
        val destPrefix = dest.absolutePath + "/"
        val res = ArrayList<Moved>()
        val files = ArrayList<File>()
        fun collect(d: File) {
            val list = d.listFiles() ?: return
            for (c in list) {
                if (c.name.startsWith(".")) continue
                val ap = c.absolutePath
                if (ap.startsWith(destPrefix) || ap + "/" == destPrefix) continue
                if (c.isDirectory) {
                    if (isBundle(c)) files.add(c) else collect(c)
                } else files.add(c)
            }
        }
        collect(src)
        for ((i, f) in files.withIndex()) {
            onProgress(i, files.size, f.name)
            val t = unique(dest, f.name)
            if (moveFile(f, t)) res.add(Moved(f, t))
        }
        onProgress(files.size, files.size, "Khali folders hata raha hai…")
        val empties = deleteEmptyDirs(src)
        return res to empties
    }

    fun deleteEmptyDirs(dir: File): Int {
        var n = 0
        val subs = dir.listFiles() ?: return 0
        for (d in subs) {
            if (!d.isDirectory) continue
            n += deleteEmptyDirs(d)
            if (d.list()?.isEmpty() == true && d.delete()) n++
        }
        return n
    }

    /** .tmp files, thumbnail cache, empty folders. Returns (freed bytes, empty folders deleted). */
    fun cleanJunk(): Pair<Long, Int> {
        var freed = 0L
        fun walk(d: File, top: Boolean) {
            val list = d.listFiles() ?: return
            for (f in list) {
                if (f.isDirectory) {
                    if (top && f.name == "Android") continue
                    if (f.name == ".thumbnails") {
                        freed += f.walkTopDown().filter { it.isFile }.sumOf { it.length() }
                        f.deleteRecursively()
                    } else walk(f, false)
                } else {
                    val e = f.extension.lowercase()
                    if (e == "tmp" || e == "temp" || f.name.startsWith(".thumbdata")) {
                        freed += f.length()
                        f.delete()
                    }
                }
            }
        }
        walk(root, true)
        val empties = deleteEmptyDirs(root)
        return freed to empties
    }

    // ---------- Duplicates ----------
    private fun hash(f: File, limit: Long = Long.MAX_VALUE): String? = try {
        val md = MessageDigest.getInstance("SHA-1")
        val buf = ByteArray(65536)
        var left = limit
        f.inputStream().use { s ->
            while (left > 0) {
                val r = s.read(buf, 0, minOf(buf.size.toLong(), left).toInt())
                if (r <= 0) break
                md.update(buf, 0, r)
                left -= r
            }
        }
        md.digest().joinToString("") { "%02x".format(it) }
    } catch (e: Exception) { null }

    /** Fast mode: same size + same date (din). Content check nahi hota. */
    fun findDuplicatesFast(items: List<Item>): List<List<Item>> {
        val fmt = SimpleDateFormat("yyyyMMdd", Locale.US)
        return items.filter { it.size > 0 }
            .groupBy { "${it.size}|${fmt.format(Date(it.mtime))}" }
            .values.filter { it.size > 1 }
            .map { g -> g.sortedBy { it.mtime } }
            .sortedByDescending { it[0].size * it.size }
    }

    /** Size match -> pehle 256KB hash -> phir full hash. Har group oldest-first. */
    fun findDuplicates(items: List<Item>): List<List<Item>> {
        val out = ArrayList<List<Item>>()
        for (g in items.filter { it.size > 0 }.groupBy { it.size }.values) {
            if (g.size < 2) continue
            for (q in g.groupBy { hash(it.file, 262144) }.filterKeys { it != null }.values) {
                if (q.size < 2) continue
                for (full in q.groupBy { hash(it.file) }.filterKeys { it != null }.values) {
                    if (full.size > 1) out.add(full.sortedBy { it.mtime })
                }
            }
        }
        return out.sortedByDescending { it[0].size * it.size }
    }
}

// ---------- Vault (hide + PIN) ----------
object Vault {
    val dir: File get() = File(Fm.root, ".FUVault")
    private val manifest: File get() = File(dir, "manifest.txt")

    fun entries(): List<Pair<String, String>> {
        if (!manifest.exists()) return emptyList()
        return manifest.readLines().mapNotNull {
            val p = it.split('\t')
            if (p.size == 2) p[0] to p[1] else null
        }.filter { File(dir, it.first).exists() }
    }

    private fun write(list: List<Pair<String, String>>) {
        dir.mkdirs()
        manifest.writeText(list.joinToString("") { "${it.first}\t${it.second}\n" })
    }

    fun hide(files: List<File>): Int {
        dir.mkdirs()
        File(dir, ".nomedia").createNewFile()
        val cur = entries().toMutableList()
        var n = 0
        for (f in files) {
            val t = Fm.unique(dir, f.name)
            if (Fm.moveFile(f, t)) { cur.add(t.name to f.absolutePath); n++ }
        }
        write(cur)
        return n
    }

    fun restore(names: List<String>): Int {
        val cur = entries().toMutableList()
        var n = 0
        for (nm in names) {
            val e = cur.firstOrNull { it.first == nm } ?: continue
            val orig = File(e.second)
            val t = Fm.unique(orig.parentFile ?: Fm.root, orig.name)
            if (Fm.moveFile(File(dir, nm), t)) { cur.remove(e); n++ }
        }
        write(cur)
        return n
    }
}

object Pin {
    private fun sp(c: Context) = c.getSharedPreferences("fu", Context.MODE_PRIVATE)
    private fun h(s: String) = MessageDigest.getInstance("SHA-256")
        .digest(("fu$s").toByteArray()).joinToString("") { "%02x".format(it) }
    fun isSet(c: Context) = sp(c).contains("pin")
    fun set(c: Context, p: String) = sp(c).edit().putString("pin", h(p)).apply()
    fun check(c: Context, p: String) = sp(c).getString("pin", null) == h(p)
}
