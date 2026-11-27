package com.fu.xplayer

import android.annotation.SuppressLint
import android.content.Context
import android.content.Intent
import android.os.Build
import android.webkit.URLUtil
import androidx.compose.runtime.*
import androidx.core.content.ContextCompat
import kotlinx.coroutines.*
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.io.IOException
import java.io.InputStream
import java.io.RandomAccessFile
import java.net.HttpURLConnection
import java.net.URL
import java.net.URLDecoder
import javax.crypto.Cipher
import javax.crypto.spec.IvParameterSpec
import javax.crypto.spec.SecretKeySpec
import kotlin.coroutines.resume

/** App settings. */
object Prefs {
    private lateinit var sp: android.content.SharedPreferences
    var connections by mutableIntStateOf(8)
    var maxActive by mutableIntStateOf(3)
    var hlsToMp4 by mutableStateOf(true)
    var showHidden by mutableStateOf(false)
    var bgPlay by mutableStateOf(true)
    var adBlock by mutableStateOf(true)
    var desktopMode by mutableStateOf(false)
    var homePage by mutableStateOf("https://www.google.com")

    fun init(ctx: Context) {
        sp = ctx.getSharedPreferences("fu", Context.MODE_PRIVATE)
        connections = sp.getInt("conn", 8); maxActive = sp.getInt("max", 3)
        hlsToMp4 = sp.getBoolean("hls", true); showHidden = sp.getBoolean("hidden", false)
        bgPlay = sp.getBoolean("bg", true); adBlock = sp.getBoolean("ad", true)
        desktopMode = sp.getBoolean("desk", false); homePage = sp.getString("home", homePage) ?: homePage
    }

    fun save() {
        sp.edit().putInt("conn", connections).putInt("max", maxActive).putBoolean("hls", hlsToMp4)
            .putBoolean("hidden", showHidden).putBoolean("bg", bgPlay).putBoolean("ad", adBlock)
            .putBoolean("desk", desktopMode).putString("home", homePage).apply()
    }

    fun getPos(key: String) = sp.getLong("pos:$key", 0L)
    fun setPos(key: String, v: Long) = sp.edit().putLong("pos:$key", v).apply()
}

enum class DState { QUEUED, RUNNING, PAUSED, DONE, FAILED, CONVERTING }

class Task(
    val id: Long,
    val url: String,
    name: String,
    dir: String,
    val headers: Map<String, String>,
    var autoName: Boolean
) {
    var name by mutableStateOf(name)
    var dir by mutableStateOf(dir)
    var state by mutableStateOf(DState.QUEUED)
    var total by mutableLongStateOf(-1L)
    var done by mutableLongStateOf(0L)
    var speed by mutableLongStateOf(0L)
    var segDone by mutableIntStateOf(0)
    var segTotal by mutableIntStateOf(0)
    var error by mutableStateOf("")
    var conns by mutableIntStateOf(1)
    var added: Long = System.currentTimeMillis()
    @Volatile var job: Job? = null
    /** multi-part state: [start, end, pos] per part */
    var parts: Array<LongArray>? = null

    val file get() = File(dir, name)
    val temp get() = File(dir, "$name.fudl")
    val isHls get() = url.substringBefore('?').lowercase().endsWith(".m3u8") || url.contains(".m3u8?")
    val fraction: Float?
        get() = when {
            state == DState.DONE -> 1f
            segTotal > 0 -> segDone.toFloat() / segTotal
            total > 0 -> done.toFloat() / total
            else -> null
        }
}

object Dl {
    val tasks = mutableStateListOf<Task>()
    private val scope = CoroutineScope(SupervisorJob() + Dispatchers.IO)
    @SuppressLint("StaticFieldLeak")
    lateinit var app: Context
    private var loaded = false
    const val UA = "Mozilla/5.0 (Linux; Android 13) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Mobile Safari/537.36"

    val base: File get() = File(Fs.downloads, "FU XPlayer")

    fun categoryDir(name: String): File = File(
        base, when (Fs.kindOfExt(name.substringAfterLast('.', "").lowercase())) {
            Kind.VIDEO -> "Video"; Kind.AUDIO -> "Music"; Kind.IMAGE -> "Pictures"; Kind.DOC -> "Documents"
            Kind.ARCHIVE -> "Compressed"; Kind.APK -> "Programs"; else -> "Other"
        }
    )

    fun init(ctx: Context) {
        if (loaded) return
        app = ctx.applicationContext; loaded = true
        load()
    }

    fun active() = tasks.count { it.state == DState.RUNNING || it.state == DState.CONVERTING }

    fun guessName(url: String, cd: String? = null, mime: String? = null): String {
        var n = try { URLUtil.guessFileName(url, cd, mime) } catch (_: Exception) { "download" }
        n = try { URLDecoder.decode(n, "UTF-8") } catch (_: Exception) { n }
        if (url.substringBefore('?').lowercase().endsWith(".m3u8")) {
            n = n.substringBeforeLast('.') .ifBlank { "video" }
            if (n == "index" || n == "master" || n == "playlist" || n == "chunklist") n = "video_" + System.currentTimeMillis() / 1000
            n += ".mp4"
        }
        return clean(n)
    }

    fun clean(n: String) = n.replace(Regex("[\\\\/:*?\"<>|\\n\\r]"), "_").take(150).ifBlank { "download" }

    fun add(url: String, name: String? = null, headers: Map<String, String> = emptyMap(), startNow: Boolean = true): Task {
        val nm = clean(name?.takeIf { it.isNotBlank() } ?: guessName(url))
        val dir = categoryDir(nm)
        val t = Task(System.nanoTime(), url.trim(), nm, dir.absolutePath, headers, name.isNullOrBlank())
        if (!startNow) t.state = DState.PAUSED
        tasks.add(0, t)
        save(); pump()
        return t
    }

    @Synchronized
    fun pump() {
        var running = active()
        for (t in tasks.reversed()) {
            if (running >= Prefs.maxActive) break
            if (t.state == DState.QUEUED) { start(t); running++ }
        }
    }

    private fun start(t: Task) {
        t.state = DState.RUNNING; t.error = ""
        startService()
        t.job = scope.launch { run(t) }
    }

    fun pause(t: Task) {
        if (t.state == DState.RUNNING || t.state == DState.QUEUED) {
            t.state = DState.PAUSED; t.job?.cancel(); t.speed = 0; save()
        }
    }

    fun resume(t: Task) {
        if (t.state == DState.PAUSED || t.state == DState.FAILED) { t.state = DState.QUEUED; pump() }
    }

    fun restart(t: Task) {
        t.job?.cancel(); t.temp.delete(); t.parts = null; t.done = 0; t.total = -1; t.segDone = 0; t.segTotal = 0
        hlsDir(t).deleteRecursively()
        t.state = DState.QUEUED; pump()
    }

    fun remove(t: Task, deleteFile: Boolean) {
        t.job?.cancel(); tasks.remove(t)
        t.temp.delete(); hlsDir(t).deleteRecursively()
        if (deleteFile) t.file.delete()
        save(); pump()
    }

    fun pauseAll() = tasks.forEach { pause(it) }
    fun resumeAll() = tasks.forEach { resume(it) }
    fun clearDone() { tasks.removeAll { it.state == DState.DONE }; save() }

    private fun startService() {
        try {
            val i = Intent(app, DlService::class.java)
            if (Build.VERSION.SDK_INT >= 26) ContextCompat.startForegroundService(app, i) else app.startService(i)
        } catch (_: Exception) {}
    }

    // -------------- worker --------------
    private suspend fun run(t: Task) {
        val ticker = scope.launch {
            var last = t.done
            while (isActive) { delay(1000); t.speed = (t.done - last).coerceAtLeast(0); last = t.done }
        }
        try {
            File(t.dir).mkdirs()
            if (t.isHls) hls(t) else http(t)
            t.state = DState.DONE; t.speed = 0
            Fs.scan(app, listOf(t.file.absolutePath))
            notifyDone(t, true)
        } catch (e: CancellationException) {
            if (t.state == DState.RUNNING) t.state = DState.PAUSED
        } catch (e: Throwable) {
            t.state = DState.FAILED; t.error = e.message ?: e.javaClass.simpleName
            notifyDone(t, false)
        } finally {
            ticker.cancel(); t.speed = 0
            save()
            withContext(NonCancellable + Dispatchers.Main) { pump() }
        }
    }

    private fun open(url: String, t: Task, range: String? = null): HttpURLConnection {
        var u = url
        repeat(8) {
            val c = URL(u).openConnection() as HttpURLConnection
            c.connectTimeout = 20000; c.readTimeout = 30000
            c.instanceFollowRedirects = false
            c.setRequestProperty("User-Agent", t.headers["User-Agent"] ?: UA)
            c.setRequestProperty("Accept", "*/*")
            c.setRequestProperty("Accept-Encoding", "identity")
            t.headers.forEach { (k, v) -> if (k != "User-Agent" && v.isNotBlank()) c.setRequestProperty(k, v) }
            if (range != null) c.setRequestProperty("Range", range)
            val code = c.responseCode
            if (code in 300..399) {
                val loc = c.getHeaderField("Location") ?: throw IOException("Redirect without location")
                u = URL(URL(u), loc).toString(); c.disconnect(); return@repeat
            }
            if (code >= 400) { c.disconnect(); throw IOException("HTTP $code") }
            return c
        }
        throw IOException("Too many redirects")
    }

    private suspend fun http(t: Task) {
        val probe = open(t.url, t, "bytes=0-")
        val ranged = probe.responseCode == 206
        val cr = probe.getHeaderField("Content-Range")
        val len = cr?.substringAfter('/')?.toLongOrNull() ?: probe.contentLengthLong
        if (t.autoName && t.done == 0L) {
            val n = clean(guessName(probe.url.toString(), probe.getHeaderField("Content-Disposition"), probe.contentType?.substringBefore(';')))
            if (n != t.name) {
                t.name = Fs.unique(categoryDir(n).also { it.mkdirs() }, n).name
                t.dir = categoryDir(n).absolutePath
            }
            t.autoName = false
        } else if (t.done == 0L && t.file.exists()) {
            t.name = Fs.unique(File(t.dir), t.name).name
        }
        File(t.dir).mkdirs()
        t.total = len
        val conns = Prefs.connections.coerceIn(1, 16)
        if (ranged && len > 2L * 1024 * 1024 && conns > 1) {
            probe.disconnect()
            multi(t, len, conns)
        } else {
            single(t, probe, ranged, len)
        }
        if (!t.temp.renameTo(t.file)) {
            t.temp.copyTo(t.file, true); t.temp.delete()
        }
    }

    private suspend fun single(t: Task, probe: HttpURLConnection, ranged: Boolean, len: Long) {
        t.conns = 1
        val existing = if (t.temp.exists()) t.temp.length() else 0L
        val conn: HttpURLConnection
        val append: Boolean
        if (existing > 0 && ranged) {
            probe.disconnect()
            conn = open(t.url, t, "bytes=$existing-"); append = conn.responseCode == 206
        } else { conn = probe; append = false }
        t.done = if (append) existing else 0
        RandomAccessFile(t.temp, "rw").use { raf ->
            if (!append) raf.setLength(0)
            raf.seek(t.done)
            conn.inputStream.use { pumpStream(it, raf, t) { n -> t.done += n } }
        }
        conn.disconnect()
        if (len > 0 && t.done < len) throw IOException("Connection toot gaya (${fmtSize(t.done)}/${fmtSize(len)}) — Resume dabayen")
    }

    private inline fun pumpStream(i: InputStream, raf: RandomAccessFile, t: Task, onRead: (Int) -> Unit) {
        val buf = ByteArray(64 * 1024)
        while (true) {
            if (t.state != DState.RUNNING) throw CancellationException("paused")
            val n = i.read(buf); if (n < 0) break
            raf.write(buf, 0, n); onRead(n)
        }
    }

    private suspend fun multi(t: Task, len: Long, conns: Int) = coroutineScope {
        val parts = t.parts?.takeIf { t.temp.exists() } ?: run {
            val n = conns.coerceAtMost((len / (1024 * 1024)).toInt().coerceAtLeast(1))
            val size = len / n
            Array(n) { i ->
                val s = i * size; val e = if (i == n - 1) len - 1 else s + size - 1
                longArrayOf(s, e, s)
            }.also { t.parts = it }
        }
        t.conns = parts.size
        RandomAccessFile(t.temp, "rw").use { if (it.length() != len) it.setLength(len) }
        fun sum() = parts.sumOf { it[2] - it[0] }
        t.done = sum()
        parts.filter { it[2] <= it[1] }.map { p ->
            async {
                var tries = 0
                while (p[2] <= p[1]) {
                    try {
                        val c = open(t.url, t, "bytes=${p[2]}-${p[1]}")
                        if (c.responseCode != 206) { c.disconnect(); throw IOException("Server ne range support band kar di") }
                        RandomAccessFile(t.temp, "rw").use { raf ->
                            raf.seek(p[2])
                            c.inputStream.use { i ->
                                val buf = ByteArray(64 * 1024)
                                while (p[2] <= p[1]) {
                                    if (t.state != DState.RUNNING) throw CancellationException("paused")
                                    val want = minOf(buf.size.toLong(), p[1] - p[2] + 1).toInt()
                                    val n = i.read(buf, 0, want); if (n < 0) break
                                    raf.write(buf, 0, n); p[2] += n
                                    synchronized(t) { t.done = sum() }
                                }
                            }
                        }
                        c.disconnect()
                        if (p[2] <= p[1]) throw IOException("part incomplete")
                    } catch (e: CancellationException) { throw e } catch (e: IOException) {
                        if (++tries > 6) throw e
                        delay(1500L * tries)
                    }
                }
            }
        }.awaitAll()
        t.done = len
    }

    // -------------- HLS (.m3u8) --------------
    private fun hlsDir(t: Task) = File(app.cacheDir, "hls_${t.id}")

    private fun fetchText(url: String, t: Task): String {
        val c = open(url, t); try { return c.inputStream.bufferedReader().readText() } finally { c.disconnect() }
    }

    private fun fetchBytes(url: String, t: Task): ByteArray {
        val c = open(url, t); try { return c.inputStream.readBytes() } finally { c.disconnect() }
    }

    private fun attr(line: String, key: String): String? =
        Regex("$key=(\"[^\"]*\"|[^,]*)").find(line)?.groupValues?.get(1)?.trim('"')

    private class Seg(val url: String, val keyUrl: String?, val iv: ByteArray?, val seq: Long)

    private suspend fun hls(t: Task) = coroutineScope {
        var url = t.url
        var text = fetchText(url, t)
        if (text.contains("#EXT-X-STREAM-INF")) {
            val lines = text.lines()
            var best = -1L; var bestUrl: String? = null
            for (i in lines.indices) if (lines[i].startsWith("#EXT-X-STREAM-INF")) {
                val bw = attr(lines[i], "BANDWIDTH")?.toLongOrNull() ?: 0
                val u = lines.drop(i + 1).firstOrNull { it.isNotBlank() && !it.startsWith("#") } ?: continue
                if (bw > best) { best = bw; bestUrl = URL(URL(url), u.trim()).toString() }
            }
            url = bestUrl ?: throw IOException("Playlist mai stream nahi mili")
            text = fetchText(url, t)
        }
        if (text.contains("METHOD=SAMPLE-AES")) throw IOException("Ye video DRM (SAMPLE-AES) protected hai, download nahi ho sakti")
        val segs = ArrayList<Seg>()
        var keyUrl: String? = null; var iv: ByteArray? = null
        var seq = Regex("#EXT-X-MEDIA-SEQUENCE:(\\d+)").find(text)?.groupValues?.get(1)?.toLongOrNull() ?: 0L
        var initUrl: String? = null
        for (raw in text.lines()) {
            val l = raw.trim()
            when {
                l.startsWith("#EXT-X-KEY") -> {
                    val m = attr(l, "METHOD")
                    if (m == "NONE") { keyUrl = null; iv = null } else {
                        keyUrl = attr(l, "URI")?.let { URL(URL(url), it).toString() }
                        iv = attr(l, "IV")?.removePrefix("0x")?.removePrefix("0X")?.let { hex ->
                            hex.padStart(32, '0').chunked(2).map { it.toInt(16).toByte() }.toByteArray()
                        }
                    }
                }
                l.startsWith("#EXT-X-MAP") -> initUrl = attr(l, "URI")?.let { URL(URL(url), it).toString() }
                l.isNotEmpty() && !l.startsWith("#") -> { segs += Seg(URL(URL(url), l).toString(), keyUrl, iv, seq); seq++ }
            }
        }
        if (segs.isEmpty()) throw IOException("Playlist khali hai")
        val dir = hlsDir(t).also { it.mkdirs() }
        t.segTotal = segs.size; t.conns = Prefs.connections.coerceIn(1, 8)
        val keys = HashMap<String, ByteArray>()
        fun key(u: String) = synchronized(keys) { keys.getOrPut(u) { fetchBytes(u, t) } }
        fun segFile(i: Int) = File(dir, "%06d.seg".format(i))
        t.segDone = segs.indices.count { segFile(it).exists() }
        t.done = segs.indices.sumOf { segFile(it).let { f -> if (f.exists()) f.length() else 0L } }
        val sem = kotlinx.coroutines.sync.Semaphore(t.conns)
        segs.mapIndexed { i, s ->
            async {
                if (segFile(i).exists()) return@async
                sem.acquire()
                try {
                    var tries = 0
                    while (true) {
                        if (t.state != DState.RUNNING) throw CancellationException("paused")
                        try {
                            var b = fetchBytes(s.url, t)
                            if (s.keyUrl != null) {
                                val ivb = s.iv ?: ByteArray(16).also { a -> var v = s.seq; for (k in 15 downTo 8) { a[k] = (v and 0xff).toByte(); v = v shr 8 } }
                                val c = Cipher.getInstance("AES/CBC/PKCS5Padding")
                                c.init(Cipher.DECRYPT_MODE, SecretKeySpec(key(s.keyUrl), "AES"), IvParameterSpec(ivb))
                                b = c.doFinal(b)
                            }
                            val tmp = File(dir, "%06d.part".format(i)); tmp.writeBytes(b); tmp.renameTo(segFile(i))
                            synchronized(t) { t.segDone++; t.done += b.size }
                            break
                        } catch (e: IOException) {
                            if (++tries > 5) throw e
                            delay(1000L * tries)
                        }
                    }
                } finally { sem.release() }
            }
        }.awaitAll()
        // join segments
        val fmp4 = initUrl != null
        val joined = File(t.dir, t.name.substringBeforeLast('.') + if (fmp4) ".mp4" else ".ts")
        val out = Fs.unique(File(t.dir), joined.name)
        out.outputStream().buffered().use { o ->
            initUrl?.let { o.write(fetchBytes(it, t)) }
            for (i in segs.indices) segFile(i).inputStream().use { it.copyTo(o) }
        }
        dir.deleteRecursively()
        if (!fmp4 && Prefs.hlsToMp4) {
            t.state = DState.CONVERTING
            val mp4 = Fs.unique(File(t.dir), t.name.substringBeforeLast('.') + ".mp4")
            val ok = withContext(Dispatchers.Main) {
                suspendCancellableCoroutine { cont ->
                    Convert.toMp4(app, out, mp4, { p -> t.segDone = p; t.segTotal = 100 }) { ok, err ->
                        if (!ok) t.error = "MP4 convert fail ($err) — .ts file rakh li"
                        if (cont.isActive) cont.resume(ok)
                    }
                }
            }
            if (ok) { out.delete(); t.name = mp4.name } else t.name = out.name
        } else t.name = out.name
        t.segDone = t.segTotal
        t.total = t.file.length(); t.done = t.total
    }

    /** "Download complete" / "failed" notification; tap opens the file. */
    fun notifyDone(t: Task, ok: Boolean) = Notify.show(
        app, (t.id % 100000).toInt() + 100,
        if (ok) "✔ Download complete" else "✖ Download fail",
        if (ok) "${t.name}  (${fmtSize(t.file.length())})" else "${t.name}: ${t.error}",
        if (ok) t.file else null
    )

    // -------------- persistence --------------
    private val store get() = File(app.filesDir, "downloads.json")

    fun save() {
        try {
            val a = JSONArray()
            tasks.toList().forEach { t ->
                a.put(JSONObject().apply {
                    put("id", t.id); put("url", t.url); put("name", t.name); put("dir", t.dir)
                    put("state", t.state.name); put("total", t.total); put("done", t.done)
                    put("auto", t.autoName); put("err", t.error); put("added", t.added)
                    put("h", JSONObject(t.headers))
                    t.parts?.let { ps -> put("parts", JSONArray(ps.map { JSONArray(it.toList()) })) }
                })
            }
            store.writeText(a.toString())
        } catch (_: Exception) {}
    }

    private fun load() {
        try {
            if (!store.exists()) return
            val a = JSONArray(store.readText())
            for (i in 0 until a.length()) {
                val o = a.getJSONObject(i)
                val h = o.optJSONObject("h")
                val hm = HashMap<String, String>()
                h?.keys()?.forEach { k -> hm[k] = h.getString(k) }
                val t = Task(o.getLong("id"), o.getString("url"), o.getString("name"), o.getString("dir"), hm, o.optBoolean("auto"))
                val st = DState.valueOf(o.getString("state"))
                t.state = if (st == DState.RUNNING || st == DState.QUEUED || st == DState.CONVERTING) DState.PAUSED else st
                t.total = o.optLong("total", -1); t.done = o.optLong("done", 0); t.error = o.optString("err")
                t.added = o.optLong("added")
                o.optJSONArray("parts")?.let { pa ->
                    t.parts = Array(pa.length()) { j -> pa.getJSONArray(j).let { longArrayOf(it.getLong(0), it.getLong(1), it.getLong(2)) } }
                }
                tasks.add(t)
            }
        } catch (_: Exception) {}
    }
}
