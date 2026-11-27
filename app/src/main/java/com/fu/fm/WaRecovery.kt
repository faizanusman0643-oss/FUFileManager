package com.fu.fm

import android.content.Context
import java.io.File

/**
 * WhatsApp deleted media recovery.
 *
 * Root ke baghair Android pe delete hui file wapas nahi aa sakti, is liye tareeqa ye hai:
 * WhatsApp ke Media folders ki har nayi file ki copy foran `.FURecovery` mai bana lo
 * (WaService live dekhta rehta hai). Jab WhatsApp/user original delete kare to copy yahan bachi rehti hai.
 * Saath mai Android 11+ ke trash (`.trashed-*`) wali WhatsApp files bhi dikhate hain.
 */
data class Rec(val item: Item, val original: File, val deleted: Boolean, val status: Boolean, val trashed: Boolean)

object WaRecovery {
    /** label -> WhatsApp Media folder (storage root se relative). Label hi backup ka subfolder hai. */
    private val SRC = listOf(
        "WhatsApp" to "Android/media/com.whatsapp/WhatsApp/Media",
        "WhatsApp-old" to "WhatsApp/Media",
        "Business" to "Android/media/com.whatsapp.w4b/WhatsApp Business/Media",
        "Business-old" to "WhatsApp Business/Media"
    )
    private val SKIP_DIRS = setOf("WhatsApp Stickers", "WallPaper")

    val backup: File get() = File(Fm.root, ".FURecovery")
    val outDir: File get() = File(Fm.root, "Download/WhatsApp Recovered")

    private fun srcDir(label: String) = File(Fm.root, SRC.first { it.first == label }.second)

    /** (label, folder) jo phone pe maujood hain. */
    fun sources(): List<Pair<String, File>> =
        SRC.map { it.first to File(Fm.root, it.second) }.filter { it.second.isDirectory }

    fun wantedDir(d: File) = !(d.name.startsWith(".") && d.name != ".Statuses") && d.name !in SKIP_DIRS
    private fun wantedFile(f: File) = !f.name.startsWith(".") && f.length() > 0

    // ---------- Settings ----------
    private fun sp(c: Context) = c.getSharedPreferences("fu", Context.MODE_PRIVATE)
    fun enabled(c: Context) = sp(c).getBoolean("wa_auto", false)
    fun setEnabled(c: Context, on: Boolean) = sp(c).edit().putBoolean("wa_auto", on).apply()

    // ---------- Backup ----------
    /** WhatsApp file ka backup path, ya null agar ye file backup ke laiq nahi. */
    fun backupFor(f: File): File? {
        val ap = f.absolutePath
        for ((label, dir) in sources()) {
            val pre = dir.absolutePath + "/"
            if (!ap.startsWith(pre)) continue
            val rel = ap.removePrefix(pre)
            val parts = rel.split('/')
            if (parts.dropLast(1).any { !wantedDir(File(it)) }) return null
            return File(File(backup, label), rel)
        }
        return null
    }

    /** Ek file ki copy banao. true = nayi copy bani. */
    @Synchronized
    fun save(f: File): Boolean {
        if (!f.isFile || !wantedFile(f)) return false
        val t = backupFor(f) ?: return false
        if (t.exists() && t.length() == f.length()) return false
        return try {
            backup.mkdirs()
            File(backup, ".nomedia").let { if (!it.exists()) it.createNewFile() }
            t.parentFile?.mkdirs()
            val tmp = File(t.parentFile, ".${t.name}.part")
            f.copyTo(tmp, true)
            t.delete()
            if (!tmp.renameTo(t)) { tmp.delete(); false }
            else { t.setLastModified(f.lastModified()); true }
        } catch (e: Exception) { false }
    }

    /** Saare WhatsApp media ka backup update karo. Returns nayi copies ki ginti. */
    fun sync(onProgress: (Int, Int, String) -> Unit = { _, _, _ -> }): Int {
        val files = ArrayList<File>()
        fun walk(d: File) {
            for (c in d.listFiles() ?: return) {
                if (c.isDirectory) { if (wantedDir(c)) walk(c) } else if (wantedFile(c)) files.add(c)
            }
        }
        sources().forEach { walk(it.second) }
        var n = 0
        for ((i, f) in files.withIndex()) {
            onProgress(i, files.size, f.name)
            if (save(f)) n++
        }
        onProgress(files.size, files.size, "")
        return n
    }

    /** Sab backup + trash wali WhatsApp files. Deleted wali pehle, naye se purane. */
    fun list(): List<Rec> {
        val out = ArrayList<Rec>()
        for ((label, _) in SRC) {
            val base = File(backup, label)
            if (!base.isDirectory) continue
            val pre = base.absolutePath + "/"
            base.walkTopDown().filter { it.isFile && !it.name.startsWith(".") }.forEach { b ->
                val rel = b.absolutePath.removePrefix(pre)
                val orig = File(srcDir(label), rel)
                out.add(Rec(b.toItem(), orig, !orig.exists(), rel.startsWith(".Statuses/"), false))
            }
        }
        // Android 11+ trash: ".trashed-<expiry>-<naam>"
        val rx = Regex("^\\.trashed-\\d+-(.+)$")
        for ((_, dir) in sources()) {
            dir.walkTopDown().onEnter { it == dir || wantedDir(it) }.filter { it.isFile }.forEach { f ->
                val m = rx.find(f.name) ?: return@forEach
                val name = m.groupValues[1]
                val item = Item(f.absolutePath, name, f.length(), f.lastModified(), catOf(name))
                out.add(Rec(item, File(f.parentFile, name), true, false, true))
            }
        }
        return out.sortedWith(compareBy<Rec> { !it.deleted }.thenByDescending { it.item.mtime })
    }

    /** Chuni hui files `Download/WhatsApp Recovered/<Photos|Videos|…>` mai copy. Returns naye paths. */
    fun recover(recs: List<Rec>, onProgress: (Int, Int, String) -> Unit = { _, _, _ -> }): List<String> {
        val done = ArrayList<String>()
        for ((i, r) in recs.withIndex()) {
            onProgress(i, recs.size, r.item.name)
            val src = r.item.file
            if (!src.isFile) continue
            val t = Fm.unique(File(outDir, r.item.cat.label), r.item.name)
            if (Fm.copyAny(src, t)) { t.setLastModified(src.lastModified()); done.add(t.absolutePath) }
        }
        onProgress(recs.size, recs.size, "")
        return done
    }

    /** Wo backup copies hatao jin ki original abhi phone mai hai aur jo `days` din se purani hain. */
    fun prune(days: Int): Pair<Int, Long> {
        val cutoff = System.currentTimeMillis() - days * 86_400_000L
        var n = 0; var b = 0L
        list().filter { !it.trashed && !it.deleted && it.item.mtime < cutoff }.forEach {
            val s = it.item.size
            if (it.item.file.delete()) { n++; b += s }
        }
        Fm.deleteEmptyDirs(backup)
        return n to b
    }

    fun backupSize(recs: List<Rec>) = recs.filter { !it.trashed }.sumOf { it.item.size }
}
