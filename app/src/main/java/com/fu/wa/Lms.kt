package com.fu.wa

import android.content.Context
import android.text.Html
import androidx.work.Constraints
import androidx.work.ExistingPeriodicWorkPolicy
import androidx.work.NetworkType
import androidx.work.PeriodicWorkRequestBuilder
import androidx.work.WorkManager
import androidx.work.Worker
import androidx.work.WorkerParameters
import java.net.HttpURLConnection
import java.net.URL
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import java.util.Locale
import java.util.concurrent.TimeUnit

/** Reads the student's AIOU (Aaghi LMS) course pages and turns workshops / quizzes into alarms. */
object Lms {
    const val BASE = "https://lms.aiou.edu.pk"
    private const val HOST = "lms.aiou.edu.pk"

    class Expired : Exception("AIOU login expire ho gaya — app mein dobara 'AIOU Login' karein")

    data class Resp(val code: Int, val body: String, val location: String?, val url: String)

    private val teamsRx = Regex("""https://teams\.(?:microsoft|live)\.com/[^"'\s<>]+""")

    fun http(c: Context, url: String, follow: Boolean, depth: Int = 0): Resp {
        val conn = URL(url).openConnection() as HttpURLConnection
        try {
            conn.instanceFollowRedirects = false
            conn.connectTimeout = 12000
            conn.readTimeout = 15000
            val ck = Store.cookie(c)
            if (ck.isNotBlank()) conn.setRequestProperty("Cookie", ck)
            conn.setRequestProperty("User-Agent", Store.ua(c))
            conn.setRequestProperty("Accept-Language", "en-US,en;q=0.9")
            val code = conn.responseCode
            for ((k, v) in conn.headerFields) {
                if (k != null && k.equals("Set-Cookie", true)) v.forEach { Store.mergeCookie(c, it) }
            }
            val loc = conn.getHeaderField("Location")?.let { URL(URL(url), it).toString() }
            if (code in 300..399 && loc != null) {
                if (follow && depth < 6 && URL(loc).host.equals(HOST, true) && !loc.contains("/login/")) {
                    return http(c, loc, true, depth + 1)
                }
                return Resp(code, "", loc, url)
            }
            val stream = if (code >= 400) conn.errorStream else conn.inputStream
            val body = stream?.bufferedReader()?.use { it.readText() } ?: ""
            return Resp(code, body, null, url)
        } finally {
            conn.disconnect()
        }
    }

    private fun expired(r: Resp): Boolean {
        val l = r.location ?: ""
        return l.contains("microsoftonline", true) || l.contains("/login/", true) ||
            r.body.contains("page-login-index") || r.code == 401 || r.code == 403
    }

    fun text(html: String): String =
        Html.fromHtml(html.replace(Regex("(?is)<(script|style)[^>]*>.*?</(script|style)>"), " "),
            Html.FROM_HTML_MODE_LEGACY).toString()
            .replace(' ', ' ').replace(Regex("\\s+"), " ").trim()

    private fun dt(d: String, mon: String, y: String, h: String, mi: String, ap: String): LocalDateTime? {
        val m = Parse.mon(mon)
        if (m <= 0) return null
        var hh = h.toIntOrNull() ?: return null
        val pm = ap.uppercase().startsWith("P")
        if (pm && hh != 12) hh += 12
        if (!pm && hh == 12) hh = 0
        return runCatching { LocalDateTime.of(y.toInt(), m, d.toInt(), hh, mi.toInt()) }.getOrNull()
    }

    private val nameRx = Regex("""(\d{1,2})-([A-Za-z]{3,9})-(\d{4})\s*(\d{1,2}):(\d{2})\s*([AaPp][Mm])""")

    private fun after(t: String, label: String): LocalDateTime? {
        val rx = Regex("""(?i)(?:$label)\s*:?\s*(?:[A-Za-z]+,\s*)?(\d{1,2})\s+([A-Za-z]+)\s+(\d{4}),?\s*(\d{1,2}):(\d{2})\s*([AP]M)""")
        val g = rx.find(t)?.groupValues ?: return null
        return dt(g[1], g[2], g[3], g[4], g[5], g[6])
    }

    private fun hm(d: LocalDateTime) = "%02d:%02d".format(d.hour, d.minute)

    private fun titleCase(s: String) = s.lowercase().split(" ").joinToString(" ") { w ->
        w.replaceFirstChar { it.uppercase() }
    }

    fun parseCourse(html: String, courseId: String): List<Workshop> {
        val title = Regex("(?is)<title>(.*?)</title>").find(html)?.groupValues?.get(1)?.let { text(it) } ?: ""
        val tokens = title.removePrefix("Course:").split("|").map { it.trim() }
        val code = tokens.firstOrNull { it.matches(Regex("\\d{3,5}")) }
            ?: Regex("""WORKSHOP\d*\|(\d{3,5})\|""", RegexOption.IGNORE_CASE).find(html)?.groupValues?.get(1)
            ?: ""
        val cname = tokens.firstOrNull { t ->
            t.length > 4 && t.contains(' ') && t.any { it.isLetter() } &&
                !t.contains("AAGHI", true) && !t.contains("WORKSHOP", true) && !t.startsWith("Course", true)
        } ?: ""
        val label = listOf(code, titleCase(cname)).filter { it.isNotBlank() }.joinToString(" ").ifBlank { "Course $courseId" }
        val courseUrl = "$BASE/course/view.php?id=$courseId"

        val out = mutableListOf<Workshop>()
        val ms = Regex("""<li[^>]*?\bid="module-(\d+)"[^>]*>""").findAll(html).toList()
        for ((i, m) in ms.withIndex()) {
            val end = if (i + 1 < ms.size) ms[i + 1].range.first else minOf(html.length, m.range.first + 30000)
            val chunk = html.substring(m.range.first, end)
            val cmid = m.groupValues[1]
            val type = Regex("modtype_(\\w+)").find(chunk)?.groupValues?.get(1) ?: ""
            val rawName = Regex("""data-activityname="([^"]*)"""").find(chunk)?.groupValues?.get(1)
                ?: Regex("""(?s)class="instancename[^"]*"[^>]*>(.*?)<""").find(chunk)?.groupValues?.get(1)
                ?: ""
            val name = text(rawName)
            val t = text(chunk)
            val locked = t.contains("Available from", true) || t.contains("Not available", true)
            val teams = teamsRx.find(chunk)?.value?.replace("&amp;", "&") ?: ""
            val modUrl = "$BASE/mod/$type/view.php?id=$cmid"
            val nm = nameRx.find(name)

            if (type == "quiz") {
                val open = after(t, "Opens|Opened|Available from") ?: continue
                val close = after(t, "Closes|Closed|Due")
                out += Workshop(
                    id = cmid.toLong(), course = label, title = name.ifBlank { "Quiz" },
                    date = open.toLocalDate().toString(), start = hm(open),
                    quizDate = close?.toLocalDate()?.toString() ?: "",
                    quizTime = close?.let { hm(it) } ?: "",
                    kind = "quiz", auto = true, courseId = courseId,
                    lmsUrl = courseUrl, modUrl = modUrl, locked = locked, notes = name,
                )
            } else if (name.contains("workshop", true) || (type == "url" && nm != null)) {
                val at = nm?.groupValues?.let { g -> dt(g[1], g[2], g[3], g[4], g[5], g[6]) }
                    ?: after(t, "Available from") ?: continue
                val num = Regex("(?i)workshop\\s*0*(\\d+)").find(name)?.groupValues?.get(1)
                out += Workshop(
                    id = cmid.toLong(), course = label,
                    title = if (num != null) "Workshop $num" else name,
                    date = at.toLocalDate().toString(), start = hm(at), link = teams,
                    kind = "workshop", auto = true, courseId = courseId,
                    lmsUrl = courseUrl, modUrl = modUrl, locked = locked, notes = name,
                )
            }
        }
        return out
    }

    /** Opens an unlocked Moodle URL activity and returns the real Teams link behind it. */
    fun resolve(c: Context, modUrl: String): String {
        if (modUrl.isBlank()) return ""
        val r = http(c, "$modUrl&redirect=1", false)
        val l = r.location
        if (l != null) {
            val host = runCatching { URL(l).host }.getOrDefault("")
            if (!host.equals(HOST, true) && !host.contains("microsoftonline", true)) return l
        }
        teamsRx.find(r.body)?.let { return it.value.replace("&amp;", "&") }
        Regex("""(?s)class="urlworkaround".*?href="([^"]+)"""").find(r.body)?.let {
            return it.groupValues[1].replace("&amp;", "&")
        }
        return ""
    }

    @Synchronized
    fun syncCourses(c: Context, ids: List<String>, notifyNew: Boolean = true): String {
        if (ids.isEmpty()) return "Koi course nahi — pehle 'AIOU Login' karein"
        if (Store.cookie(c).isBlank()) return "Pehle 'AIOU Login' karein"
        val old = Store.load(c)
        val oldMap = old.associateBy { it.id }
        val byCourse = LinkedHashMap<String, List<Workshop>>()
        var err = ""
        for (id in ids) {
            try {
                val r = http(c, "$BASE/course/view.php?id=$id", true)
                if (expired(r)) throw Expired()
                byCourse[id] = parseCourse(r.body, id)
            } catch (e: Expired) {
                val was = Store.loginOk(c)
                Store.setLoginOk(c, false)
                Store.setLastSync(c, "⚠ " + e.message)
                if (was) Notify.show(c, "AIOU login dobara karein",
                    "Session expire ho gaya — FU Workshop Alert khol kar 'AIOU Login' dabayein", "", 777001, false)
                return e.message ?: "Login expire"
            } catch (e: Exception) {
                err = e.message ?: e.toString()
            }
        }
        if (byCourse.isNotEmpty()) Store.setLoginOk(c, true)
        val today = LocalDate.now()
        val merged = mutableListOf<Workshop>()
        for ((_, list) in byCourse) {
            for (n0 in list) {
                var n = n0
                val o = oldMap[n.id]
                val upcoming = n.startDate?.isBefore(today) != true
                if (n.kind == "workshop" && n.link.isBlank() && !n.locked && upcoming) {
                    val l = runCatching { resolve(c, n.modUrl) }.getOrDefault("")
                    if (l.isNotBlank()) n = n.copy(link = l)
                }
                if (n.link.isBlank() && o != null && o.link.isNotBlank()) n = n.copy(link = o.link)
                n = n.copy(done = o?.done ?: false)
                if (notifyNew && n.kind == "workshop" && upcoming && n.link.isNotBlank() && o?.link.isNullOrBlank()) {
                    Notify.show(c, "Teams link aa gaya ✔",
                        "${n.label} — ${fmtDate(n.startDate)}, ${fmtTime(parseTime(n.start))}",
                        n.link, ("L" + n.id).hashCode() and 0x7fffffff, false)
                }
                merged += n
            }
        }
        val replaced = byCourse.filterValues { it.isNotEmpty() }.keys
        val keep = old.filter { !(it.auto && it.courseId in replaced) }
        Store.save(c, (keep + merged).sortedWith(compareBy({ it.done }, { it.date }, { it.start })))
        Alarms.rescheduleAll(c)
        val nW = merged.count { it.kind == "workshop" }
        val nQ = merged.count { it.kind == "quiz" }
        val msg = when {
            byCourse.isEmpty() -> "Sync nahi hua: $err"
            merged.isEmpty() -> "Courses khule lekin workshops nahi mile"
            else -> "✔ ${replaced.size} courses · $nW workshops · $nQ quiz"
        } + if (err.isNotBlank() && byCourse.isNotEmpty()) " (kuch error: $err)" else ""
        val now = LocalDateTime.now().format(DateTimeFormatter.ofPattern("d MMM, h:mm a", Locale.ENGLISH))
        Store.setLastSync(c, "$now — $msg")
        return msg
    }

    fun syncAll(c: Context) = syncCourses(c, Store.courses(c))
}

class SyncWorker(ctx: Context, p: WorkerParameters) : Worker(ctx, p) {
    override fun doWork(): Result {
        runCatching { Lms.syncAll(applicationContext) }
        return Result.success()
    }

    companion object {
        fun schedule(c: Context) {
            val req = PeriodicWorkRequestBuilder<SyncWorker>(1, TimeUnit.HOURS)
                .setConstraints(Constraints.Builder().setRequiredNetworkType(NetworkType.CONNECTED).build())
                .build()
            WorkManager.getInstance(c).enqueueUniquePeriodicWork("fuwa_sync", ExistingPeriodicWorkPolicy.KEEP, req)
        }
    }
}
