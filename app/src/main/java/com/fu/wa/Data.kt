package com.fu.wa

import android.content.Context
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.time.LocalDate
import java.time.LocalTime
import java.time.format.DateTimeFormatter
import java.util.Locale

const val AAGHI = "https://aaghi.aiou.edu.pk"

data class Workshop(
    val id: Long = System.currentTimeMillis(),
    val course: String = "",
    val title: String = "",
    val date: String = "",
    val days: Int = 1,
    val start: String = "10:00",
    val end: String = "",
    val link: String = "",
    val quizDate: String = "",
    val quizTime: String = "09:00",
    val quizEnd: String = "",
    val notes: String = "",
    val done: Boolean = false,
) {
    val startDate: LocalDate? get() = parseDate(date)
    val lastDate: LocalDate? get() = startDate?.plusDays((days.coerceAtLeast(1) - 1).toLong())
    val label: String
        get() = listOf(course, title).filter { it.isNotBlank() }.joinToString(" · ").ifBlank { "Workshop" }

    fun toJson(): JSONObject = JSONObject().apply {
        put("id", id); put("course", course); put("title", title); put("date", date)
        put("days", days); put("start", start); put("end", end); put("link", link)
        put("quizDate", quizDate); put("quizTime", quizTime); put("quizEnd", quizEnd)
        put("notes", notes); put("done", done)
    }

    companion object {
        fun from(o: JSONObject) = Workshop(
            id = o.optLong("id", System.currentTimeMillis()),
            course = o.optString("course"),
            title = o.optString("title"),
            date = o.optString("date"),
            days = o.optInt("days", 1),
            start = o.optString("start", "10:00"),
            end = o.optString("end"),
            link = o.optString("link"),
            quizDate = o.optString("quizDate"),
            quizTime = o.optString("quizTime", "09:00"),
            quizEnd = o.optString("quizEnd"),
            notes = o.optString("notes"),
            done = o.optBoolean("done", false),
        )
    }
}

fun parseDate(s: String?): LocalDate? =
    if (s.isNullOrBlank()) null else runCatching { LocalDate.parse(s.trim()) }.getOrNull()

fun parseTime(s: String?): LocalTime? =
    if (s.isNullOrBlank()) null else runCatching { LocalTime.parse(s.trim()) }.getOrNull()

private val dFmt = DateTimeFormatter.ofPattern("EEE, d MMM yyyy", Locale.ENGLISH)
private val tFmt = DateTimeFormatter.ofPattern("h:mm a", Locale.ENGLISH)
fun fmtDate(d: LocalDate?): String = d?.format(dFmt) ?: "—"
fun fmtTime(t: LocalTime?): String = t?.format(tFmt) ?: "—"

object Store {
    private fun sp(c: Context) = c.getSharedPreferences("fuwa", Context.MODE_PRIVATE)

    fun load(c: Context): List<Workshop> {
        val raw = sp(c).getString("workshops", null) ?: return emptyList()
        return importJson(raw)
    }

    fun save(c: Context, list: List<Workshop>) {
        val json = exportJson(list)
        sp(c).edit().putString("workshops", json).apply()
        runCatching {
            val dir = c.getExternalFilesDir("backup") ?: return@runCatching
            File(dir, "workshops.json").writeText(json)
        }
    }

    fun exportJson(list: List<Workshop>): String {
        val a = JSONArray()
        list.forEach { a.put(it.toJson()) }
        return a.toString(2)
    }

    fun importJson(s: String): List<Workshop> = runCatching {
        val a = JSONArray(s.trim())
        (0 until a.length()).map { Workshop.from(a.getJSONObject(it)) }
    }.getOrDefault(emptyList())

    fun minutesBefore(c: Context) = sp(c).getInt("minBefore", 30)
    fun setMinutesBefore(c: Context, v: Int) = sp(c).edit().putInt("minBefore", v).apply()
    fun dayBefore(c: Context) = sp(c).getBoolean("dayBefore", true)
    fun setDayBefore(c: Context, v: Boolean) = sp(c).edit().putBoolean("dayBefore", v).apply()
    fun codes(c: Context): Set<String> = sp(c).getStringSet("codes", emptySet()) ?: emptySet()
    fun setCodes(c: Context, v: Set<String>) = sp(c).edit().putStringSet("codes", v).apply()
}

/** Reads AIOU SMS / Teams invite / any pasted text and guesses workshop fields. */
object Parse {
    private val months = listOf("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
    private val urlRx = Regex("""https?://[^\s<>"']+""")

    fun link(t: String): String {
        val all = urlRx.findAll(t).map { it.value.trimEnd('.', ',', ')', ']') }.toList()
        return all.firstOrNull { it.contains("teams", true) } ?: all.firstOrNull() ?: ""
    }

    private fun mk(y: Int, m: Int, d: Int): LocalDate? = runCatching { LocalDate.of(y, m, d) }.getOrNull()
    private fun mon(s: String): Int {
        val l = s.lowercase()
        return months.indexOfFirst { l.startsWith(it) } + 1
    }

    fun date(text: String): LocalDate? {
        val t = urlRx.replace(text, " ")
        Regex("""(\d{4})-(\d{1,2})-(\d{1,2})""").find(t)?.let { g ->
            mk(g.groupValues[1].toInt(), g.groupValues[2].toInt(), g.groupValues[3].toInt())?.let { return it }
        }
        Regex("""(\d{1,2})(?:st|nd|rd|th)?[\s-]+([A-Za-z]{3,9})[,\s-]+(\d{4})""").find(t)?.let { g ->
            val m = mon(g.groupValues[2])
            if (m > 0) mk(g.groupValues[3].toInt(), m, g.groupValues[1].toInt())?.let { return it }
        }
        Regex("""([A-Za-z]{3,9})\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})""").find(t)?.let { g ->
            val m = mon(g.groupValues[1])
            if (m > 0) mk(g.groupValues[3].toInt(), m, g.groupValues[2].toInt())?.let { return it }
        }
        Regex("""\b(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})\b""").find(t)?.let { g ->
            var y = g.groupValues[3].toInt()
            if (y < 100) y += 2000
            mk(y, g.groupValues[2].toInt(), g.groupValues[1].toInt())?.let { return it }
        }
        return null
    }

    fun time(text: String): String? {
        val t = urlRx.replace(text, " ")
        Regex("""\b(\d{1,2})(?::(\d{2}))?\s*([AaPp])\.?\s*[Mm]\b""").find(t)?.let { g ->
            var h = g.groupValues[1].toInt()
            val m = g.groupValues[2].ifBlank { "0" }.toInt()
            val pm = g.groupValues[3].equals("p", true)
            if (h in 1..12 && m in 0..59) {
                if (pm && h != 12) h += 12
                if (!pm && h == 12) h = 0
                return "%02d:%02d".format(h, m)
            }
        }
        Regex("""\b([01]?\d|2[0-3]):([0-5]\d)\b""").find(t)?.let { g ->
            return "%02d:%02d".format(g.groupValues[1].toInt(), g.groupValues[2].toInt())
        }
        return null
    }

    fun course(text: String): String {
        val t = urlRx.replace(text, " ")
        Regex("""(?i)course\s*(?:code|no\.?)?\s*[:#-]?\s*(\d{3,5})""").find(t)?.let { return it.groupValues[1] }
        return ""
    }

    fun toWorkshop(text: String): Workshop {
        val d = date(text)
        return Workshop(
            course = course(text),
            title = "Workshop",
            date = d?.toString() ?: "",
            start = time(text) ?: "10:00",
            link = link(text),
            quizDate = d?.plusDays(4)?.toString() ?: "",
            notes = text.trim().take(300),
        )
    }
}
