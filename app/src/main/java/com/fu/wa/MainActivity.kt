package com.fu.wa

import android.Manifest
import android.app.DatePickerDialog
import android.app.TimePickerDialog
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.PowerManager
import android.provider.Settings
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Add
import androidx.compose.material.icons.filled.Delete
import androidx.compose.material.icons.filled.Edit
import androidx.compose.material.icons.filled.Settings
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalClipboardManager
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.delay
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.time.Duration
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.LocalTime

val Navy = Color(0xFF0B2545)
val Teal = Color(0xFF13A8A0)

fun fuColors() = lightColorScheme(
    primary = Teal, onPrimary = Color.White, secondary = Navy, onSecondary = Color.White,
    primaryContainer = Color(0xFFD7F3F1), onPrimaryContainer = Navy,
    surface = Color.White, background = Color(0xFFF3F6F9),
)

class MainActivity : ComponentActivity() {
    private var shared by mutableStateOf<String?>(null)
    private var alert by mutableStateOf<List<String>?>(null)
    private var resumeTick by mutableIntStateOf(0)

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        handle(intent)
        if (Build.VERSION.SDK_INT >= 33 &&
            checkSelfPermission(Manifest.permission.POST_NOTIFICATIONS) != PackageManager.PERMISSION_GRANTED
        ) requestPermissions(arrayOf(Manifest.permission.POST_NOTIFICATIONS), 1)
        Notify.ensure(this)
        Alarms.rescheduleAll(this)
        runCatching { SyncWorker.schedule(this) }
        setContent {
            MaterialTheme(colorScheme = fuColors()) {
                App(shared, { shared = null }, alert, { alert = null }, resumeTick)
            }
        }
    }

    override fun onResume() {
        super.onResume()
        resumeTick++
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        setIntent(intent)
        handle(intent)
    }

    private fun handle(i: Intent?) {
        if (i == null) return
        if (i.action == Intent.ACTION_SEND) {
            val t = listOfNotNull(i.getStringExtra(Intent.EXTRA_SUBJECT), i.getStringExtra(Intent.EXTRA_TEXT))
                .joinToString("\n")
            if (t.isNotBlank()) shared = t
            i.action = null
        }
        val at = i.getStringExtra("alert_title")
        if (at != null) {
            alert = listOf(at, i.getStringExtra("alert_text") ?: "", i.getStringExtra("alert_link") ?: "")
            i.removeExtra("alert_title")
            if (Build.VERSION.SDK_INT >= 27) { setShowWhenLocked(true); setTurnScreenOn(true) }
            val link = i.getStringExtra("alert_link") ?: ""
            if (i.getBooleanExtra("auto_open", false) && link.isNotBlank()) {
                i.removeExtra("auto_open")
                getSystemService(android.app.NotificationManager::class.java)?.cancelAll()
                window.decorView.postDelayed({ openUrl(this, link) }, 400)
            }
        }
    }
}

fun openUrl(c: Context, url: String) {
    if (url.isBlank()) { Toast.makeText(c, "Link nahi hai", Toast.LENGTH_SHORT).show(); return }
    runCatching { c.startActivity(Intent(Intent.ACTION_VIEW, Uri.parse(url))) }
        .onFailure { Toast.makeText(c, "Link open nahi hua", Toast.LENGTH_SHORT).show() }
}

fun pickDate(c: Context, cur: String, onPick: (String) -> Unit) {
    val d = parseDate(cur) ?: LocalDate.now()
    DatePickerDialog(c, { _, y, m, day -> onPick(LocalDate.of(y, m + 1, day).toString()) },
        d.year, d.monthValue - 1, d.dayOfMonth).show()
}

fun pickTime(c: Context, cur: String, onPick: (String) -> Unit) {
    val t = parseTime(cur) ?: LocalTime.of(10, 0)
    TimePickerDialog(c, { _, h, mi -> onPick("%02d:%02d".format(h, mi)) }, t.hour, t.minute, false).show()
}

data class Session(val w: Workshop, val day: Int, val start: LocalDateTime, val end: LocalDateTime)

fun sessions(list: List<Workshop>): List<Session> = list.filter { !it.done }.flatMap { w ->
    val d0 = w.startDate ?: return@flatMap emptyList<Session>()
    val st = parseTime(w.start) ?: LocalTime.of(10, 0)
    val en = parseTime(w.end)
    (0 until w.days.coerceIn(1, 30)).map { i ->
        val s = d0.plusDays(i.toLong()).atTime(st)
        val e = if (en != null && en.isAfter(st)) d0.plusDays(i.toLong()).atTime(en) else s.plusHours(2)
        Session(w, i + 1, s, e)
    }
}

fun countdown(to: LocalDateTime): String {
    val d = Duration.between(LocalDateTime.now(), to)
    if (d.isNegative) return "abhi"
    val days = d.toDays(); val h = d.toHours() % 24; val m = d.toMinutes() % 60
    return when {
        days > 0 -> "$days din $h ghante baad"
        h > 0 -> "$h ghante $m min baad"
        else -> "$m min baad"
    }
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun App(shared: String?, clearShared: () -> Unit, alert: List<String>?, clearAlert: () -> Unit, resumeTick: Int) {
    val ctx = LocalContext.current
    val scope = rememberCoroutineScope()
    var syncing by remember { mutableStateOf(false) }
    val clip = LocalClipboardManager.current
    var list by remember { mutableStateOf(Store.load(ctx)) }
    var editing by remember { mutableStateOf<Workshop?>(null) }
    var deleting by remember { mutableStateOf<Workshop?>(null) }
    var showSettings by remember { mutableStateOf(false) }
    var tick by remember { mutableLongStateOf(0L) }

    LaunchedEffect(Unit) { while (true) { delay(30_000); tick++ } }
    LaunchedEffect(resumeTick) { list = Store.load(ctx); tick++ }
    LaunchedEffect(shared) {
        if (shared != null) { editing = Parse.toWorkshop(shared); clearShared() }
    }

    fun commit(n: List<Workshop>) {
        list = n.sortedWith(compareBy({ it.done }, { it.date }, { it.start }))
        Store.save(ctx, list)
        Alarms.rescheduleAll(ctx)
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = { Text("FU Workshop Alert", fontWeight = FontWeight.Bold) },
                colors = TopAppBarDefaults.topAppBarColors(
                    containerColor = Navy, titleContentColor = Color.White, actionIconContentColor = Color.White),
                actions = {
                    IconButton(onClick = { showSettings = true }) { Icon(Icons.Default.Settings, "Settings") }
                }
            )
        },
        floatingActionButton = {
            ExtendedFloatingActionButton(
                onClick = { editing = Workshop() },
                containerColor = Teal, contentColor = Color.White,
                icon = { Icon(Icons.Default.Add, null) }, text = { Text("Workshop") })
        },
        containerColor = Color(0xFFF3F6F9),
    ) { pad ->
        LazyColumn(
            Modifier.padding(pad).fillMaxSize(),
            contentPadding = PaddingValues(12.dp, 12.dp, 12.dp, 96.dp),
            verticalArrangement = Arrangement.spacedBy(10.dp)
        ) {
            item(key = "next") { key(tick) { NextCard(list) } }
            item(key = "quick") {
                Row(Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                    OutlinedButton(onClick = { openUrl(ctx, AAGHI) }, Modifier.weight(1f)) { Text("Aaghi LMS") }
                    OutlinedButton(onClick = {
                        val t = clip.getText()?.text
                        if (t.isNullOrBlank()) Toast.makeText(ctx, "Pehle SMS ya link copy karein", Toast.LENGTH_SHORT).show()
                        else editing = Parse.toWorkshop(t)
                    }, Modifier.weight(1f)) { Text("Paste SMS/Link") }
                }
            }
            item(key = "sync") {
                key(tick, syncing) {
                    SyncCard(
                        syncing = syncing,
                        onLogin = { ctx.startActivity(Intent(ctx, LoginActivity::class.java)) },
                        onSync = {
                            syncing = true
                            scope.launch {
                                val m = withContext(Dispatchers.IO) {
                                    runCatching { Lms.syncAll(ctx) }.getOrElse { "Error: ${it.message}" }
                                }
                                syncing = false
                                list = Store.load(ctx)
                                Toast.makeText(ctx, m, Toast.LENGTH_LONG).show()
                            }
                        })
                }
            }
            if (list.isEmpty()) {
                item(key = "empty") {
                    Card(Modifier.fillMaxWidth()) {
                        Column(Modifier.padding(16.dp)) {
                            Text("Abhi koi workshop nahi", fontWeight = FontWeight.Bold)
                            Spacer(Modifier.height(6.dp))
                            Text("• + Workshop dabayein, ya\n• AIOU SMS / Teams link copy karke \"Paste SMS/Link\" dabayein, ya\n• Kisi bhi app se link Share karke FU Workshop Alert chunein.\n\nQuiz date khud workshop ke 5th din par set ho jati hai.")
                        }
                    }
                }
            }
            items(list, key = { it.id }) { w ->
                WorkshopCard(w,
                    onJoin = { openUrl(ctx, w.link) },
                    onLms = { openUrl(ctx, w.lmsUrl.ifBlank { AAGHI }) },
                    onEdit = { editing = w },
                    onDelete = { deleting = w },
                    onDone = { commit(list.map { if (it.id == w.id) it.copy(done = !it.done) else it }) })
            }
        }
    }

    editing?.let { w ->
        EditDialog(w, onDismiss = { editing = null }) { nw ->
            commit(list.filter { it.id != nw.id } + nw)
            editing = null
            Toast.makeText(ctx, "Save ho gaya — alarms set", Toast.LENGTH_SHORT).show()
        }
    }

    deleting?.let { w ->
        AlertDialog(onDismissRequest = { deleting = null },
            title = { Text("Delete?") }, text = { Text(w.label) },
            confirmButton = { TextButton(onClick = { commit(list.filter { it.id != w.id }); deleting = null }) { Text("Delete") } },
            dismissButton = { TextButton(onClick = { deleting = null }) { Text("Cancel") } })
    }

    if (showSettings) SettingsDialog(list, onImport = { commit(it) }) { showSettings = false; Alarms.rescheduleAll(ctx) }

    alert?.let { a ->
        AlertDialog(onDismissRequest = clearAlert,
            title = { Text(a[0], fontWeight = FontWeight.Bold) },
            text = { Text(a[1], fontSize = 17.sp) },
            confirmButton = {
                if (a[2].isNotBlank()) Button(onClick = { openUrl(ctx, a[2]); clearAlert() }) {
                    Text(if (a[2].contains("teams", true)) "Join Teams" else "Open")
                }
            },
            dismissButton = { TextButton(onClick = clearAlert) { Text("Theek hai") } })
    }
}

@Composable
fun NextCard(list: List<Workshop>) {
    val ctx = LocalContext.current
    val now = LocalDateTime.now()
    val next = sessions(list.filter { it.kind != "quiz" }).filter { it.end.isAfter(now) }.minByOrNull { it.start }
    val today = LocalDate.now()
    val quizOpen = list.filter { w ->
        if (w.kind == "quiz") {
            val o = w.startDate
            val e = parseDate(w.quizDate)
            !w.done && o != null && !today.isBefore(o) && (e == null || !today.isAfter(e))
        } else {
            val q = parseDate(w.quizDate)
            val e = parseDate(w.quizEnd) ?: w.lastDate
            !w.done && q != null && !today.isBefore(q) && (e == null || !today.isAfter(e))
        }
    }
    Card(
        Modifier.fillMaxWidth(),
        colors = CardDefaults.cardColors(containerColor = Navy, contentColor = Color.White),
        shape = RoundedCornerShape(16.dp)
    ) {
        Column(Modifier.padding(16.dp)) {
            Text("AGLA WORKSHOP", color = Teal, fontSize = 12.sp, fontWeight = FontWeight.Bold)
            Spacer(Modifier.height(4.dp))
            if (next == null) {
                Text("Koi upcoming workshop nahi", fontSize = 18.sp)
            } else {
                val live = !now.isBefore(next.start)
                Text(next.w.label + if (next.w.days > 1) "  (Day ${next.day})" else "",
                    fontSize = 20.sp, fontWeight = FontWeight.Bold)
                Text("${fmtDate(next.start.toLocalDate())} · ${fmtTime(next.start.toLocalTime())}")
                Spacer(Modifier.height(6.dp))
                Text(if (live) "🔴 Abhi chal raha hai" else "⏰ " + countdown(next.start),
                    color = Color(0xFFB8F0EC), fontWeight = FontWeight.SemiBold)
                Spacer(Modifier.height(10.dp))
                if (next.w.link.isNotBlank()) {
                    Button(onClick = { openUrl(ctx, next.w.link) },
                        colors = ButtonDefaults.buttonColors(containerColor = Teal)) { Text("Join Teams") }
                } else {
                    Text(if (next.w.auto) "🔒 Teams link abhi lock hai — khulte hi notification aayegi" else "Teams link abhi nahi",
                        fontSize = 13.sp, color = Color(0xFFB8F0EC))
                    if (next.w.lmsUrl.isNotBlank()) {
                        Spacer(Modifier.height(6.dp))
                        OutlinedButton(onClick = { openUrl(ctx, next.w.lmsUrl) }) { Text("LMS course kholein", color = Color.White) }
                    }
                }
            }
            if (quizOpen.isNotEmpty()) {
                Spacer(Modifier.height(12.dp))
                Box(Modifier.fillMaxWidth().background(Color(0xFF1C3D66), RoundedCornerShape(10.dp)).padding(10.dp)) {
                    Text("📝 Quiz open: " + quizOpen.joinToString(", ") { if (it.kind == "quiz") it.course else it.label })
                }
            }
        }
    }
}

@Composable
fun WorkshopCard(w: Workshop, onJoin: () -> Unit, onLms: () -> Unit, onEdit: () -> Unit, onDelete: () -> Unit, onDone: () -> Unit) {
    val d0 = w.startDate
    val quiz = w.kind == "quiz"
    val past = if (quiz) (parseDate(w.quizDate) ?: d0)?.isBefore(LocalDate.now()) == true
        else w.lastDate?.isBefore(LocalDate.now()) == true
    Card(Modifier.fillMaxWidth(), shape = RoundedCornerShape(14.dp),
        colors = CardDefaults.cardColors(containerColor = if (w.done || past) Color(0xFFE9EDF1) else Color.White)) {
        Column(Modifier.padding(14.dp)) {
            Row(verticalAlignment = Alignment.CenterVertically) {
                Column(Modifier.weight(1f)) {
                    if (quiz) {
                        Text("📝 " + w.course, fontWeight = FontWeight.Bold, fontSize = 17.sp, color = Navy)
                        Text(w.title, fontSize = 14.sp)
                        Text("Khulega: ${fmtDate(d0)} ${fmtTime(parseTime(w.start))}", fontSize = 14.sp, color = Color(0xFF0E7C76))
                        parseDate(w.quizDate)?.let {
                            Text("Band: ${fmtDate(it)} ${fmtTime(parseTime(w.quizTime))}", fontSize = 14.sp, color = Color(0xFFC62828))
                        }
                    } else {
                        Text(w.label, fontWeight = FontWeight.Bold, fontSize = 17.sp, color = Navy)
                        val range = if (w.days > 1) "${fmtDate(d0)} → ${fmtDate(w.lastDate)} (${w.days} din)" else fmtDate(d0)
                        Text(range, fontSize = 14.sp)
                        val tm = fmtTime(parseTime(w.start)) + (parseTime(w.end)?.let { " – " + fmtTime(it) } ?: "")
                        Text("🕒 $tm", fontSize = 14.sp)
                        parseDate(w.quizDate)?.let { q ->
                            Text("📝 Quiz: ${fmtDate(q)} ${fmtTime(parseTime(w.quizTime))}" +
                                (parseDate(w.quizEnd)?.let { " → ${fmtDate(it)}" } ?: ""),
                                fontSize = 14.sp, color = Color(0xFF0E7C76))
                        }
                        if (w.link.isBlank() && w.auto && !past) {
                            Text("🔒 Teams link ${fmtTime(parseTime(w.start))} par khulega", fontSize = 13.sp, color = Color(0xFF8A6D00))
                        }
                    }
                    if (w.auto) Text("AIOU LMS se sync", fontSize = 11.sp, color = Color.Gray)
                    if (w.done) Text("✔ Mukammal", fontSize = 13.sp, color = Color(0xFF2E7D32))
                }
                IconButton(onClick = onEdit) { Icon(Icons.Default.Edit, "Edit", tint = Navy) }
                IconButton(onClick = onDelete) { Icon(Icons.Default.Delete, "Delete", tint = Color(0xFFC62828)) }
            }
            Spacer(Modifier.height(6.dp))
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                if (w.link.isNotBlank()) Button(onClick = onJoin,
                    colors = ButtonDefaults.buttonColors(containerColor = Teal)) { Text("Join") }
                if (w.lmsUrl.isNotBlank()) OutlinedButton(onClick = onLms) { Text("LMS") }
                OutlinedButton(onClick = onDone) { Text(if (w.done) "Wapas active" else "Done") }
            }
        }
    }
}

@Composable
fun SyncCard(syncing: Boolean, onLogin: () -> Unit, onSync: () -> Unit) {
    val ctx = LocalContext.current
    val courses = Store.courses(ctx)
    val last = Store.lastSync(ctx)
    val ok = Store.loginOk(ctx) && courses.isNotEmpty()
    Card(Modifier.fillMaxWidth(), shape = RoundedCornerShape(14.dp),
        colors = CardDefaults.cardColors(containerColor = Color(0xFFD7F3F1))) {
        Column(Modifier.padding(14.dp)) {
            Text("AIOU LMS Auto-Sync", fontWeight = FontWeight.Bold, color = Navy)
            Text(
                if (courses.isEmpty()) "Ek baar AIOU login karein — saare courses, workshops, time aur Teams links khud aa jayenge."
                else (if (ok) "✔ ${courses.size} courses connected · har ghante check" else "⚠ Login dobara karein") +
                    (if (last.isNotBlank()) "\nLast: $last" else ""),
                fontSize = 13.sp, color = Navy)
            Spacer(Modifier.height(8.dp))
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                Button(onClick = onLogin, colors = ButtonDefaults.buttonColors(containerColor = Navy)) {
                    Text(if (courses.isEmpty() || !ok) "AIOU Login" else "Login / Courses")
                }
                if (courses.isNotEmpty()) OutlinedButton(onClick = onSync, enabled = !syncing) {
                    Text(if (syncing) "Sync…" else "Sync now")
                }
            }
        }
    }
}

@Composable
fun EditDialog(w: Workshop, onDismiss: () -> Unit, onSave: (Workshop) -> Unit) {
    val ctx = LocalContext.current
    val clip = LocalClipboardManager.current
    var course by remember { mutableStateOf(w.course) }
    var title by remember { mutableStateOf(w.title) }
    var date by remember { mutableStateOf(w.date) }
    var days by remember { mutableStateOf(w.days.toString()) }
    var start by remember { mutableStateOf(w.start) }
    var end by remember { mutableStateOf(w.end) }
    var link by remember { mutableStateOf(w.link) }
    var quizDate by remember { mutableStateOf(w.quizDate) }
    var quizTime by remember { mutableStateOf(w.quizTime) }
    var quizEnd by remember { mutableStateOf(w.quizEnd) }
    var notes by remember { mutableStateOf(w.notes) }

    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text("Workshop") },
        text = {
            Column(Modifier.verticalScroll(rememberScrollState()), verticalArrangement = Arrangement.spacedBy(8.dp)) {
                OutlinedTextField(course, { course = it }, label = { Text("Course code / naam (e.g. 8601)") }, singleLine = true)
                OutlinedTextField(title, { title = it }, label = { Text("Title (e.g. Workshop 1)") }, singleLine = true)
                Text("Workshop date", fontSize = 13.sp)
                OutlinedButton(onClick = {
                    pickDate(ctx, date) { nd ->
                        val oldDef = parseDate(date)?.plusDays(4)?.toString()
                        if (quizDate.isBlank() || quizDate == oldDef) quizDate = parseDate(nd)?.plusDays(4)?.toString() ?: ""
                        date = nd
                    }
                }, Modifier.fillMaxWidth()) { Text(if (date.isBlank()) "Date chunein" else fmtDate(parseDate(date))) }
                OutlinedTextField(days, { days = it.filter(Char::isDigit).take(2) },
                    label = { Text("Kitne din (lagatar)") }, singleLine = true,
                    keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number))
                Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                    Column(Modifier.weight(1f)) {
                        Text("Start", fontSize = 13.sp)
                        OutlinedButton(onClick = { pickTime(ctx, start) { start = it } }, Modifier.fillMaxWidth()) {
                            Text(fmtTime(parseTime(start)))
                        }
                    }
                    Column(Modifier.weight(1f)) {
                        Text("End (optional)", fontSize = 13.sp)
                        OutlinedButton(onClick = { pickTime(ctx, end.ifBlank { start }) { end = it } }, Modifier.fillMaxWidth()) {
                            Text(if (end.isBlank()) "—" else fmtTime(parseTime(end)))
                        }
                    }
                }
                OutlinedTextField(link, { link = it }, label = { Text("Teams link") }, maxLines = 3,
                    trailingIcon = {
                        TextButton(onClick = {
                            val t = clip.getText()?.text ?: ""
                            val l = Parse.link(t)
                            if (l.isNotBlank()) link = l
                            else Toast.makeText(ctx, "Clipboard mein link nahi", Toast.LENGTH_SHORT).show()
                        }) { Text("Paste") }
                    })
                HorizontalDivider()
                Text("Quiz (AIOU: 5th din se)", fontWeight = FontWeight.SemiBold)
                Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                    OutlinedButton(onClick = { pickDate(ctx, quizDate.ifBlank { date }) { quizDate = it } }, Modifier.weight(1f)) {
                        Text(if (quizDate.isBlank()) "Quiz date" else fmtDate(parseDate(quizDate)), fontSize = 13.sp)
                    }
                    OutlinedButton(onClick = { pickTime(ctx, quizTime) { quizTime = it } }, Modifier.weight(1f)) {
                        Text(fmtTime(parseTime(quizTime)), fontSize = 13.sp)
                    }
                }
                Row(verticalAlignment = Alignment.CenterVertically) {
                    OutlinedButton(onClick = { pickDate(ctx, quizEnd.ifBlank { quizDate.ifBlank { date } }) { quizEnd = it } },
                        Modifier.weight(1f)) {
                        Text(if (quizEnd.isBlank()) "Quiz last date (optional)" else "Last: " + fmtDate(parseDate(quizEnd)), fontSize = 13.sp)
                    }
                    if (quizEnd.isNotBlank() || quizDate.isNotBlank()) TextButton(onClick = { quizEnd = ""; quizDate = "" }) { Text("Clear") }
                }
                OutlinedTextField(notes, { notes = it }, label = { Text("Notes") }, maxLines = 4)
            }
        },
        confirmButton = {
            Button(onClick = {
                if (parseDate(date) == null) {
                    Toast.makeText(ctx, "Workshop date chunein", Toast.LENGTH_SHORT).show()
                } else {
                    onSave(w.copy(course = course.trim(), title = title.trim(), date = date,
                        days = days.toIntOrNull()?.coerceIn(1, 30) ?: 1, start = start, end = end,
                        link = link.trim(), quizDate = quizDate, quizTime = quizTime, quizEnd = quizEnd,
                        notes = notes.trim()))
                }
            }) { Text("Save") }
        },
        dismissButton = { TextButton(onClick = onDismiss) { Text("Cancel") } }
    )
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun SettingsDialog(list: List<Workshop>, onImport: (List<Workshop>) -> Unit, onDismiss: () -> Unit) {
    val ctx = LocalContext.current
    var minB by remember { mutableIntStateOf(Store.minutesBefore(ctx)) }
    var dayB by remember { mutableStateOf(Store.dayBefore(ctx)) }
    var importText by remember { mutableStateOf<String?>(null) }
    val pm = ctx.getSystemService(PowerManager::class.java)
    val batteryOk = pm.isIgnoringBatteryOptimizations(ctx.packageName)

    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text("Settings") },
        text = {
            Column(Modifier.verticalScroll(rememberScrollState()), verticalArrangement = Arrangement.spacedBy(8.dp)) {
                Text("Workshop se kitne minute pehle alarm?")
                Row(horizontalArrangement = Arrangement.spacedBy(6.dp)) {
                    listOf(10, 15, 30, 60).forEach { m ->
                        FilterChip(selected = minB == m, onClick = { minB = m; Store.setMinutesBefore(ctx, m) },
                            label = { Text("$m") })
                    }
                }
                Row(verticalAlignment = Alignment.CenterVertically) {
                    Text("Ek din pehle (raat 8 baje) yaad dilayein", Modifier.weight(1f))
                    Switch(checked = dayB, onCheckedChange = { dayB = it; Store.setDayBefore(ctx, it) })
                }
                HorizontalDivider()
                Text(if (Alarms.canExact(ctx)) "✔ Exact alarms ON" else "⚠ Exact alarms OFF — alarm thora late ho sakta hai",
                    fontSize = 13.sp)
                Text(if (batteryOk) "✔ Battery restriction OFF (background theek chalega)"
                    else "⚠ Battery optimization ON hai — band karein taake alarm miss na ho", fontSize = 13.sp)
                if (!batteryOk) OutlinedButton(onClick = {
                    runCatching {
                        ctx.startActivity(Intent(Settings.ACTION_REQUEST_IGNORE_BATTERY_OPTIMIZATIONS,
                            Uri.parse("package:" + ctx.packageName)))
                    }
                }, Modifier.fillMaxWidth()) { Text("Battery optimization band karein") }
                if (!Alarms.canExact(ctx) && Build.VERSION.SDK_INT >= 31) OutlinedButton(onClick = {
                    runCatching { ctx.startActivity(Intent(Settings.ACTION_REQUEST_SCHEDULE_EXACT_ALARM,
                        Uri.parse("package:" + ctx.packageName))) }
                }, Modifier.fillMaxWidth()) { Text("Exact alarm allow karein") }
                OutlinedButton(onClick = {
                    Alarms.test(ctx)
                    Toast.makeText(ctx, "10 second mein test alarm bajega", Toast.LENGTH_SHORT).show()
                }, Modifier.fillMaxWidth()) { Text("Test alarm (10 sec)") }
                HorizontalDivider()
                Text("Backup", fontWeight = FontWeight.SemiBold)
                OutlinedButton(onClick = {
                    val i = Intent(Intent.ACTION_SEND).apply {
                        type = "text/plain"
                        putExtra(Intent.EXTRA_SUBJECT, "FU Workshop Alert backup")
                        putExtra(Intent.EXTRA_TEXT, Store.exportJson(list))
                    }
                    ctx.startActivity(Intent.createChooser(i, "Backup bhejein"))
                }, Modifier.fillMaxWidth()) { Text("Export backup (share)") }
                if (importText == null) {
                    OutlinedButton(onClick = { importText = "" }, Modifier.fillMaxWidth()) { Text("Import backup") }
                } else {
                    OutlinedTextField(importText ?: "", { importText = it }, label = { Text("Backup JSON paste karein") }, maxLines = 5)
                    Button(onClick = {
                        val got = Store.importJson(importText ?: "")
                        if (got.isEmpty()) Toast.makeText(ctx, "Backup sahi nahi", Toast.LENGTH_SHORT).show()
                        else {
                            val ids = got.map { it.id }.toSet()
                            onImport(list.filter { it.id !in ids } + got)
                            Toast.makeText(ctx, "${got.size} workshops import", Toast.LENGTH_SHORT).show()
                            importText = null
                        }
                    }, Modifier.fillMaxWidth()) { Text("Import") }
                }
                Spacer(Modifier.height(4.dp))
                Text("FU · Prince Faizan — Innovating Tomorrow", fontSize = 11.sp, color = Color.Gray)
            }
        },
        confirmButton = { TextButton(onClick = onDismiss) { Text("Done") } }
    )
}
