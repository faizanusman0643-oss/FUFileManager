package com.fu.wa

import android.Manifest
import android.app.AlarmManager
import android.app.Notification
import android.app.NotificationChannel
import android.app.NotificationManager
import android.app.PendingIntent
import android.content.BroadcastReceiver
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.media.AudioAttributes
import android.media.RingtoneManager
import android.net.Uri
import android.os.Build
import androidx.core.app.NotificationCompat
import androidx.core.app.NotificationManagerCompat
import androidx.core.content.ContextCompat
import java.time.LocalDateTime
import java.time.LocalTime
import java.time.ZoneId

data class Spec(
    val code: Int,
    val at: Long,
    val title: String,
    val text: String,
    val link: String,
    val loud: Boolean,
    val wid: Long = 0L,
    val courseId: String = "",
    val check: Boolean = false,
)

object Alarms {
    private const val ACTION = "com.fu.wa.ALERT"

    fun specs(c: Context, list: List<Workshop>): List<Spec> {
        val now = System.currentTimeMillis()
        val minB = Store.minutesBefore(c)
        val dayBefore = Store.dayBefore(c)
        val out = mutableListOf<Spec>()
        for (w in list) {
            if (w.done) continue
            val d0 = w.startDate ?: continue
            val st = parseTime(w.start) ?: LocalTime.of(10, 0)
            val name = w.label
            val openLink = w.link.ifBlank { w.lmsUrl }

            fun add(key: String, dt: LocalDateTime, title: String, text: String, link: String, loud: Boolean,
                    check: Boolean = false) {
                val ms = dt.atZone(ZoneId.systemDefault()).toInstant().toEpochMilli()
                if (ms > now) {
                    val code = ("${w.id}|$key").hashCode() and 0x7fffffff
                    out += Spec(code, ms, title, text, link, loud, w.id, w.courseId, check)
                }
            }

            if (w.kind == "quiz") {
                add("q", d0.atTime(st), "Quiz khul gaya", "${w.course} — ${w.title}",
                    w.lmsUrl.ifBlank { AAGHI }, true)
                val cd = parseDate(w.quizDate)
                if (cd != null) {
                    val ct = parseTime(w.quizTime) ?: LocalTime.of(23, 59)
                    val close = cd.atTime(ct)
                    val warn = if (ct.isAfter(LocalTime.of(12, 0))) cd.atTime(9, 0) else close.minusHours(3)
                    add("qe", warn, "Quiz aaj band hoga", "${w.course} — ${w.title}: ${fmtTime(ct)} tak attempt karein",
                        w.lmsUrl.ifBlank { AAGHI }, true)
                }
                continue
            }

            if (dayBefore) {
                add("db", d0.minusDays(1).atTime(20, 0), "Kal workshop hai",
                    "$name — ${fmtDate(d0)}, ${fmtTime(st)}", openLink, false)
            }
            for (i in 0 until w.days.coerceIn(1, 30)) {
                val d = d0.plusDays(i.toLong())
                val dayTxt = if (w.days > 1) " (Day ${i + 1})" else ""
                if (st.isAfter(LocalTime.of(8, 30))) {
                    add("m$i", d.atTime(8, 0), "Aaj workshop hai$dayTxt",
                        "$name — ${fmtTime(st)} baje", openLink, false)
                }
                add("b$i", d.atTime(st).minusMinutes(minB.toLong()), "$minB min mein workshop$dayTxt",
                    "$name — Teams join karne ke liye tayyar ho jayein", openLink, true)
                add("s$i", d.atTime(st), "Workshop shuru$dayTxt",
                    "$name — abhi Join karein", openLink, true)
                if (w.auto) {
                    listOf(60L, 20L, 5L).forEach { m ->
                        add("c$i-$m", d.atTime(st).minusMinutes(m), "", "", "", false, check = true)
                    }
                    add("c$i+2", d.atTime(st).plusMinutes(2), "", "", "", false, check = true)
                }
            }
            val qd = parseDate(w.quizDate)
            if (qd != null) {
                val qt = parseTime(w.quizTime) ?: LocalTime.of(9, 0)
                add("q", qd.atTime(qt), "Quiz khul gaya",
                    "$name — Aaghi LMS par quiz attempt karein", AAGHI, true)
                val qe = parseDate(w.quizEnd)
                if (qe != null) {
                    add("qe", qe.atTime(9, 0), "Quiz ka aakhri din",
                        "$name — aaj quiz zaroor attempt karein", AAGHI, true)
                }
            }
        }
        return out
    }

    private fun intent(c: Context, code: Int) = Intent(c, AlertReceiver::class.java).apply {
        action = ACTION
        data = Uri.parse("fuwa://a/$code")
    }

    fun canExact(c: Context): Boolean {
        val am = c.getSystemService(AlarmManager::class.java)
        return Build.VERSION.SDK_INT < 31 || am.canScheduleExactAlarms()
    }

    fun schedule(c: Context, s: Spec) {
        val am = c.getSystemService(AlarmManager::class.java)
        val i = intent(c, s.code).apply {
            putExtra("title", s.title); putExtra("text", s.text)
            putExtra("link", s.link); putExtra("code", s.code); putExtra("loud", s.loud)
            putExtra("wid", s.wid); putExtra("courseId", s.courseId); putExtra("check", s.check)
        }
        val pi = PendingIntent.getBroadcast(c, s.code, i,
            PendingIntent.FLAG_UPDATE_CURRENT or PendingIntent.FLAG_IMMUTABLE)
        try {
            if (s.check && canExact(c)) {
                am.setExactAndAllowWhileIdle(AlarmManager.RTC_WAKEUP, s.at, pi)
            } else if (canExact(c)) {
                val show = PendingIntent.getActivity(c, 0, Intent(c, MainActivity::class.java),
                    PendingIntent.FLAG_IMMUTABLE)
                am.setAlarmClock(AlarmManager.AlarmClockInfo(s.at, show), pi)
            } else {
                am.setAndAllowWhileIdle(AlarmManager.RTC_WAKEUP, s.at, pi)
            }
        } catch (e: SecurityException) {
            am.setAndAllowWhileIdle(AlarmManager.RTC_WAKEUP, s.at, pi)
        }
    }

    private fun cancel(c: Context, code: Int) {
        val am = c.getSystemService(AlarmManager::class.java)
        val pi = PendingIntent.getBroadcast(c, code, intent(c, code),
            PendingIntent.FLAG_NO_CREATE or PendingIntent.FLAG_IMMUTABLE) ?: return
        am.cancel(pi)
        pi.cancel()
    }

    /** Cancels every old alarm and schedules fresh ones from saved data. */
    fun rescheduleAll(c: Context): Int {
        Store.codes(c).forEach { it.toIntOrNull()?.let { code -> cancel(c, code) } }
        val specs = specs(c, Store.load(c)).sortedBy { it.at }.take(400)
        specs.forEach { schedule(c, it) }
        Store.setCodes(c, specs.map { it.code.toString() }.toSet())
        return specs.size
    }

    fun test(c: Context) {
        schedule(c, Spec(424242, System.currentTimeMillis() + 10_000, "Test alarm",
            "FU Workshop Alert theek kaam kar raha hai ✔", AAGHI, true))
    }
}

object Notify {
    const val CH = "fuwa_alarm_v1"

    fun ensure(c: Context) {
        val nm = c.getSystemService(NotificationManager::class.java)
        if (nm.getNotificationChannel(CH) != null) return
        val sound = RingtoneManager.getDefaultUri(RingtoneManager.TYPE_ALARM)
            ?: RingtoneManager.getDefaultUri(RingtoneManager.TYPE_NOTIFICATION)
        val ch = NotificationChannel(CH, "Workshop alarms", NotificationManager.IMPORTANCE_HIGH).apply {
            description = "AIOU workshop aur quiz alarms"
            enableVibration(true)
            vibrationPattern = longArrayOf(0, 700, 300, 700, 300, 700)
            setSound(sound, AudioAttributes.Builder()
                .setUsage(AudioAttributes.USAGE_ALARM)
                .setContentType(AudioAttributes.CONTENT_TYPE_SONIFICATION)
                .build())
            lockscreenVisibility = Notification.VISIBILITY_PUBLIC
        }
        nm.createNotificationChannel(ch)
    }

    fun show(c: Context, title: String, text: String, link: String, code: Int, loud: Boolean) {
        ensure(c)
        val openI = Intent(c, MainActivity::class.java).apply {
            flags = Intent.FLAG_ACTIVITY_NEW_TASK or Intent.FLAG_ACTIVITY_CLEAR_TOP
            putExtra("alert_title", title); putExtra("alert_text", text); putExtra("alert_link", link)
            putExtra("auto_open", loud && link.isNotBlank())
        }
        val open = PendingIntent.getActivity(c, code, openI,
            PendingIntent.FLAG_UPDATE_CURRENT or PendingIntent.FLAG_IMMUTABLE)
        val b = NotificationCompat.Builder(c, CH)
            .setSmallIcon(R.drawable.ic_stat)
            .setContentTitle(title)
            .setContentText(text)
            .setStyle(NotificationCompat.BigTextStyle().bigText(text))
            .setPriority(NotificationCompat.PRIORITY_MAX)
            .setCategory(NotificationCompat.CATEGORY_ALARM)
            .setVisibility(NotificationCompat.VISIBILITY_PUBLIC)
            .setColor(0xFF13A8A0.toInt())
            .setAutoCancel(true)
            .setContentIntent(open)
        if (loud) b.setFullScreenIntent(open, true)
        if (link.isNotBlank()) {
            val v = Intent(Intent.ACTION_VIEW, Uri.parse(link)).addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
            val vp = PendingIntent.getActivity(c, code + 1, v, PendingIntent.FLAG_IMMUTABLE)
            b.addAction(0, if (link.contains("teams", true)) "Join Teams" else "Open", vp)
        }
        val n = b.build()
        if (loud) n.flags = n.flags or Notification.FLAG_INSISTENT
        val ok = Build.VERSION.SDK_INT < 33 ||
            ContextCompat.checkSelfPermission(c, Manifest.permission.POST_NOTIFICATIONS) == PackageManager.PERMISSION_GRANTED
        if (ok) NotificationManagerCompat.from(c).notify(code, n)
    }
}

class AlertReceiver : BroadcastReceiver() {
    override fun onReceive(context: Context, intent: Intent) {
        val title = intent.getStringExtra("title") ?: "Workshop alert"
        val text = intent.getStringExtra("text") ?: ""
        val link0 = intent.getStringExtra("link") ?: ""
        val code = intent.getIntExtra("code", 1)
        val loud = intent.getBooleanExtra("loud", true)
        val check = intent.getBooleanExtra("check", false)
        val courseId = intent.getStringExtra("courseId") ?: ""
        val wid = intent.getLongExtra("wid", 0L)
        val needFresh = courseId.isNotBlank() && (check || !link0.contains("teams", true))
        if (!needFresh) {
            Notify.show(context, title, text, link0, code, loud)
            return
        }
        val pr = goAsync()
        Thread {
            try {
                runCatching { Lms.syncCourses(context, listOf(courseId), notifyNew = check) }
                if (!check) {
                    val w = Store.load(context).firstOrNull { it.id == wid }
                    val link = w?.link?.takeIf { it.isNotBlank() } ?: link0
                    Notify.show(context, title, text, link, code, loud)
                }
            } catch (e: Exception) {
                if (!check) Notify.show(context, title, text, link0, code, loud)
            } finally {
                pr.finish()
            }
        }.start()
    }
}

class BootReceiver : BroadcastReceiver() {
    override fun onReceive(context: Context, intent: Intent) {
        Alarms.rescheduleAll(context)
    }
}
