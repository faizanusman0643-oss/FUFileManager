package com.fu.xplayer

import android.Manifest
import android.content.ClipboardManager
import android.content.ComponentName
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.provider.Settings
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.compose.foundation.*
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import androidx.media3.session.MediaController
import androidx.media3.session.SessionToken
import coil.Coil
import coil.ImageLoader
import coil.decode.VideoFrameDecoder
import java.io.File

enum class Screen { DESKTOP, EXPLORER, PLAYER, DOWNLOADS, BROWSER }

object UiState {
    var front by mutableStateOf(Screen.DESKTOP)
    val open = mutableStateListOf<Screen>()
    var addUrl by mutableStateOf<Triple<String, String?, Map<String, String>>?>(null)
    var settings by mutableStateOf(false)
    fun show(s: Screen) { if (s != Screen.DESKTOP && s !in open) open.add(s); front = s }
    fun close(s: Screen) { open.remove(s); front = open.lastOrNull() ?: Screen.DESKTOP }
}

class MainActivity : ComponentActivity() {
    private var granted by mutableStateOf(false)
    private var controller: com.google.common.util.concurrent.ListenableFuture<MediaController>? = null

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        Prefs.init(this); Dl.init(this); DlService.channel(this)
        Coil.setImageLoader(ImageLoader.Builder(this).components { add(VideoFrameDecoder.Factory()) }.crossfade(true).build())
        if (Build.VERSION.SDK_INT >= 33 && ContextCompat.checkSelfPermission(this, Manifest.permission.POST_NOTIFICATIONS) != PackageManager.PERMISSION_GRANTED)
            ActivityCompat.requestPermissions(this, arrayOf(Manifest.permission.POST_NOTIFICATIONS), 2)
        handleIntent(intent)
        setContent {
            W7Theme {
                if (granted) Shell { finishAndRemoveTask() } else PermissionScreen { requestAccess() }
            }
        }
    }

    override fun onStart() {
        super.onStart()
        try { controller = MediaController.Builder(this, SessionToken(this, ComponentName(this, PlaybackService::class.java))).buildAsync() } catch (_: Exception) {}
    }

    override fun onStop() {
        super.onStop()
        val p = Players.get(this)
        if (!Prefs.bgPlay && !isChangingConfigurations) p.pause()
        controller?.let { MediaController.releaseFuture(it) }; controller = null
    }

    override fun onNewIntent(intent: Intent) { super.onNewIntent(intent); handleIntent(intent) }

    private fun handleIntent(i: Intent?) {
        i ?: return
        i.getStringExtra("screen")?.let { if (it == "downloads") UiState.show(Screen.DOWNLOADS) else if (it == "player") UiState.show(Screen.PLAYER) }
        i.getStringExtra("open")?.let { path ->
            val f = File(path)
            if (Fs.isMedia(f)) { Queue.play(listOf(f.toPItem()), 0); UiState.show(Screen.PLAYER) }
            else { f.parentFile?.let { ExplorerNav.request = Place.Dir(it) }; UiState.show(Screen.EXPLORER) }
        }
        when (i.action) {
            Intent.ACTION_VIEW -> i.data?.let { u ->
                val type = i.type ?: contentResolver.getType(u) ?: ""
                val s = u.toString()
                if (u.scheme?.startsWith("http") == true && !type.startsWith("video") && !type.startsWith("audio") && !Regex("\\.(mp4|mkv|mp3|m3u8|webm|m4a|avi|mov)(\\?|$)", RegexOption.IGNORE_CASE).containsMatchIn(s)) {
                    BrowserState.pending = s; UiState.show(Screen.BROWSER)
                } else {
                    Queue.play(listOf(PItem(u, u.lastPathSegment?.substringAfterLast('/') ?: "Media", type.startsWith("audio"))), 0)
                    UiState.show(Screen.PLAYER)
                }
            }
            Intent.ACTION_SEND -> {
                val text = i.getStringExtra(Intent.EXTRA_TEXT) ?: ""
                Regex("https?://\\S+").find(text)?.value?.let { UiState.addUrl = Triple(it, null, emptyMap()) }
            }
        }
    }

    override fun onResume() { super.onResume(); granted = hasAccess() }

    private fun hasAccess(): Boolean =
        if (Build.VERSION.SDK_INT >= 30) Environment.isExternalStorageManager()
        else ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED

    private fun requestAccess() {
        if (Build.VERSION.SDK_INT >= 30) {
            try { startActivity(Intent(Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION, Uri.parse("package:$packageName"))) }
            catch (e: Exception) { startActivity(Intent(Settings.ACTION_MANAGE_ALL_FILES_ACCESS_PERMISSION)) }
        } else ActivityCompat.requestPermissions(this, arrayOf(Manifest.permission.READ_EXTERNAL_STORAGE, Manifest.permission.WRITE_EXTERNAL_STORAGE), 1)
    }

    override fun onRequestPermissionsResult(requestCode: Int, permissions: Array<String>, grantResults: IntArray) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults); granted = hasAccess()
    }
}

@Composable
fun PermissionScreen(onGrant: () -> Unit) {
    Box(Modifier.fillMaxSize().background(W7.desk), contentAlignment = Alignment.Center) {
        Column(Modifier.padding(24.dp).background(W7.aero).border(1.dp, Color(0xFF5A7FA8)).padding(4.dp)) {
            Text("FU XPlayer", Modifier.padding(6.dp), fontSize = 13.sp)
            Column(Modifier.background(Color.White).padding(18.dp)) {
                Text("Welcome — Innovating Tomorrow", fontSize = 18.sp, color = Color(0xFF1E3A5F))
                Spacer(Modifier.height(8.dp))
                Text("File manager, player aur downloads ke liye 'All files access' chahiye. Neeche button dabao aur FU XPlayer ko allow karo.", fontSize = 14.sp)
                Spacer(Modifier.height(16.dp))
                WinButton("Allow access", default = true, onClick = onGrant)
            }
        }
    }
}

@Composable
fun Shell(onShutdown: () -> Unit) {
    val ctx = LocalContext.current
    var start by remember { mutableStateOf(false) }
    val play: (List<PItem>, Int) -> Unit = { l, i -> Queue.play(l, i); UiState.show(Screen.PLAYER) }
    val dl: (String, String?, Map<String, String>) -> Unit = { u, n, h -> UiState.addUrl = Triple(u, n, h) }

    Column(Modifier.fillMaxSize().background(W7.desk).systemBarsPadding()) {
        Box(Modifier.weight(1f).fillMaxWidth()) {
            Desktop()
            when (UiState.front) {
                Screen.EXPLORER -> ExplorerScreen({ UiState.front = Screen.DESKTOP }, { UiState.close(Screen.EXPLORER) }, play)
                Screen.PLAYER -> PlayerScreen(onClose = {
                    Players.get(ctx).let { it.stop(); it.clearMediaItems() }; Queue.items = emptyList(); UiState.close(Screen.PLAYER)
                }, onMin = { UiState.front = Screen.DESKTOP })
                Screen.DOWNLOADS -> DownloadsScreen({ UiState.front = Screen.DESKTOP }, { UiState.close(Screen.DOWNLOADS) }, play)
                Screen.BROWSER -> BrowserScreen({ UiState.front = Screen.DESKTOP }, { UiState.close(Screen.BROWSER) }, dl) { play(listOf(it), 0) }
                else -> {}
            }
            if (start) StartMenu(
                left = listOf(
                    StartEntry("FU Media Player", Icons.Filled.PlayCircle, Color(0xFFE88A1A)) { UiState.show(Screen.PLAYER) },
                    StartEntry("Windows Explorer", Icons.Filled.FolderOpen, W7.folder) { ExplorerNav.request = Place.Computer; UiState.show(Screen.EXPLORER) },
                    StartEntry("Download Manager", Icons.Filled.Download, W7.teal) { UiState.show(Screen.DOWNLOADS) },
                    StartEntry("FU Browser", Icons.Filled.Public, Color(0xFF1E88E5)) { UiState.show(Screen.BROWSER) },
                    StartEntry("Add download URL", Icons.Filled.AddLink, Color(0xFF2E7D32)) { UiState.addUrl = Triple(clip(ctx), null, emptyMap()) },
                    StartEntry("Recycle Bin", Icons.Filled.Delete, Color(0xFF607D8B)) { ExplorerNav.request = Place.Bin; UiState.show(Screen.EXPLORER) },
                    StartEntry("Settings", Icons.Filled.Settings, Color(0xFF607D8B)) { UiState.settings = true },
                ),
                right = listOf(
                    StartEntry("Faizan", Icons.Filled.Person, Color.White) { ExplorerNav.request = Place.Dir(Fs.root); UiState.show(Screen.EXPLORER) },
                    StartEntry("Documents", Icons.Filled.Description, Color.White) { ExplorerNav.request = Place.Lib(Kind.DOC); UiState.show(Screen.EXPLORER) },
                    StartEntry("Pictures", Icons.Filled.Image, Color.White) { ExplorerNav.request = Place.Lib(Kind.IMAGE); UiState.show(Screen.EXPLORER) },
                    StartEntry("Music", Icons.Filled.MusicNote, Color.White) { ExplorerNav.request = Place.Lib(Kind.AUDIO); UiState.show(Screen.EXPLORER) },
                    StartEntry("Videos", Icons.Filled.Movie, Color.White) { ExplorerNav.request = Place.Lib(Kind.VIDEO); UiState.show(Screen.EXPLORER) },
                    StartEntry("Downloads", Icons.Filled.Download, Color.White) { ExplorerNav.request = Place.Dir(Fs.downloads); UiState.show(Screen.EXPLORER) },
                    StartEntry("Computer", Icons.Filled.Computer, Color.White) { ExplorerNav.request = Place.Computer; UiState.show(Screen.EXPLORER) },
                    StartEntry("Control Panel", Icons.Filled.Settings, Color.White) { UiState.settings = true },
                ),
                onSearch = { q -> if (q.isNotBlank()) { ExplorerNav.request = Place.Search(Fs.root, q); UiState.show(Screen.EXPLORER) } },
                onShutdown = { start = false; onShutdown() },
                onDismiss = { start = false }
            )
        }
        val apps = buildList {
            add(TaskApp(Screen.EXPLORER, "Explorer", Icons.Filled.FolderOpen, W7.folder))
            add(TaskApp(Screen.PLAYER, "Player", Icons.Filled.PlayCircle, Color(0xFFFFA726)))
            add(TaskApp(Screen.DOWNLOADS, "Downloads", Icons.Filled.Download, Color(0xFF4DD0C8)))
            add(TaskApp(Screen.BROWSER, "Browser", Icons.Filled.Public, Color(0xFF64B5F6)))
        }
        Taskbar(apps, UiState.front, { start = !start }, { s ->
            if (UiState.front == s) UiState.front = Screen.DESKTOP else UiState.show(s)
        }) { UiState.front = Screen.DESKTOP }
    }

    UiState.addUrl?.let { (u, n, h) -> AddDownloadDialog(u, n, h) { UiState.addUrl = null } }
    if (UiState.settings) SettingsDialog { UiState.settings = false }
}

fun clip(ctx: Context): String = try {
    val c = ctx.getSystemService(ClipboardManager::class.java).primaryClip?.getItemAt(0)?.text?.toString() ?: ""
    if (c.startsWith("http")) c.trim() else ""
} catch (_: Exception) { "" }

@Composable
private fun Desktop() {
    Box(Modifier.fillMaxSize()) {
        // FU watermark like a Win7 wallpaper logo
        Column(Modifier.align(Alignment.Center), horizontalAlignment = Alignment.CenterHorizontally) {
            Text("FU", color = Color(0x33FFFFFF), fontSize = 120.sp, fontWeight = FontWeight.Bold)
            Text("Prince Faizan • Innovating Tomorrow", color = Color(0x55FFFFFF), fontSize = 14.sp)
        }
        Column(Modifier.padding(8.dp).verticalScroll(rememberScrollState())) {
            DesktopIcon("Computer", Icons.Filled.Computer, Color(0xFFDDEEFF)) { ExplorerNav.request = Place.Computer; UiState.show(Screen.EXPLORER) }
            DesktopIcon("Recycle Bin", Icons.Filled.Delete, Color(0xFFDDEEFF)) { ExplorerNav.request = Place.Bin; UiState.show(Screen.EXPLORER) }
            DesktopIcon("FU Media Player", Icons.Filled.PlayCircle, Color(0xFFFFA726)) { UiState.show(Screen.PLAYER) }
            DesktopIcon("Download Manager", Icons.Filled.Download, Color(0xFF4DD0C8)) { UiState.show(Screen.DOWNLOADS) }
            DesktopIcon("FU Browser", Icons.Filled.Public, Color(0xFF64B5F6)) { UiState.show(Screen.BROWSER) }
            DesktopIcon("Videos", Icons.Filled.Movie, Color(0xFFCE93D8)) { ExplorerNav.request = Place.Lib(Kind.VIDEO); UiState.show(Screen.EXPLORER) }
            DesktopIcon("Music", Icons.Filled.MusicNote, Color(0xFFFF8A80)) { ExplorerNav.request = Place.Lib(Kind.AUDIO); UiState.show(Screen.EXPLORER) }
        }
    }
}

@Composable
fun AddDownloadDialog(url0: String, name0: String?, headers: Map<String, String>, onDismiss: () -> Unit) {
    val ctx = LocalContext.current
    var url by remember { mutableStateOf(url0) }
    var name by remember { mutableStateOf(name0 ?: if (url0.isNotBlank()) Dl.guessName(url0) else "") }
    LaunchedEffect(url) { if (name0 == null && url.startsWith("http")) name = Dl.guessName(url) }
    WinDialog("Add New Download", onDismiss, Icons.Filled.Download, W7.teal, buttons = {
        WinButton("Start Download", default = true, enabled = url.startsWith("http")) {
            Dl.add(url, name.ifBlank { null }, headers); onDismiss(); UiState.show(Screen.DOWNLOADS)
        }
        WinButton("Later") { Dl.add(url, name.ifBlank { null }, headers, startNow = false); onDismiss() }
        WinButton("Cancel", onClick = onDismiss)
    }) {
        Text("Address (URL):", fontSize = 13.sp)
        Row(verticalAlignment = Alignment.CenterVertically) {
            WinField(url, { url = it }, Modifier.weight(1f), "https://...")
            IconButton({ clip(ctx).takeIf { it.isNotEmpty() }?.let { url = it } }) { Icon(Icons.Filled.ContentPaste, "Paste") }
        }
        Spacer(Modifier.height(8.dp))
        Text("Save as:", fontSize = 13.sp)
        WinField(name, { name = it })
        Spacer(Modifier.height(6.dp))
        Text("Folder: Download/FU XPlayer/${Dl.categoryDir(name.ifBlank { "x" }).name}", fontSize = 12.sp, color = W7.gray)
        if (url.contains(".m3u8")) Text("HLS stream — parts download ho ke ${if (Prefs.hlsToMp4) "MP4 ban jayegi" else ".ts file banegi"}", fontSize = 12.sp, color = W7.teal)
        if (url.startsWith("http") && !Regex("\\.[a-z0-9]{2,5}(\\?|$)", RegexOption.IGNORE_CASE).containsMatchIn(url.substringAfterLast('/')))
            Text("Ye page link lagta hai. Video page ho to 'FU Browser' mai kholen — wahan video chalte hi green Download button aa jata hai.", fontSize = 12.sp, color = Color(0xFFB06000))
        if (url.startsWith("http")) TextButton({ BrowserState.pending = url; onDismiss(); UiState.show(Screen.BROWSER) }) { Text("Browser mai kholo") }
    }
}

@Composable
fun DownloadsScreen(onMin: () -> Unit, onClose: () -> Unit, onPlay: (List<PItem>, Int) -> Unit) {
    val ctx = LocalContext.current
    var filter by remember { mutableStateOf("All") }
    val cats = listOf("All", "Running", "Finished", "Video", "Music", "Documents", "Compressed", "Programs", "Other")
    WinWindow("Download Manager", Icons.Filled.Download, W7.teal, onMin = onMin, onClose = onClose) {
        Row(Modifier.fillMaxWidth().background(W7.cmd).horizontalScroll(rememberScrollState()), verticalAlignment = Alignment.CenterVertically) {
            CmdItem("Add URL", Icons.Filled.AddLink, Color(0xFF2E7D32)) { UiState.addUrl = Triple(clip(ctx), null, emptyMap()) }
            CmdItem("Resume all", Icons.Filled.PlayArrow) { Dl.resumeAll() }
            CmdItem("Stop all", Icons.Filled.Pause) { Dl.pauseAll() }
            CmdItem("Clear finished", Icons.Filled.DeleteSweep) { Dl.clearDone() }
            CmdItem("Folder", Icons.Filled.FolderOpen, W7.folder) { Dl.base.mkdirs(); ExplorerNav.request = Place.Dir(Dl.base); UiState.show(Screen.EXPLORER) }
            CmdItem("Options", Icons.Filled.Settings) { UiState.settings = true }
        }
        Row(Modifier.fillMaxWidth().background(W7.navBg).horizontalScroll(rememberScrollState()).padding(4.dp)) {
            cats.forEach { c ->
                Text(c, fontSize = 13.sp, color = if (c == filter) Color.White else W7.text, modifier = Modifier.padding(2.dp)
                    .background(if (c == filter) W7.link else Color.Transparent).clickable { filter = c }.padding(horizontal = 10.dp, vertical = 5.dp))
            }
        }
        val list = Dl.tasks.filter { t ->
            when (filter) {
                "All" -> true
                "Running" -> t.state != DState.DONE
                "Finished" -> t.state == DState.DONE
                else -> File(t.dir).name == filter || (filter == "Music" && File(t.dir).name == "Music")
            }
        }
        if (list.isEmpty()) Text(
            "Koi download nahi.\n\n• 'Add URL' se koi bhi link (video, mp3, apk, zip, pdf...) daalen\n• Ya FU Browser mai video chalayen — green Download button aa jayega\n• Ya kisi bhi app se link 'Share' karke FU XPlayer chunen",
            color = W7.gray, modifier = Modifier.padding(20.dp), fontSize = 14.sp
        )
        LazyColumn(Modifier.weight(1f).fillMaxWidth()) {
            items(list, key = { it.id }) { t -> TaskRow(t, onPlay) }
        }
        val act = Dl.tasks.filter { it.state == DState.RUNNING }
        Text("  ${Dl.tasks.size} downloads  •  ${act.size} chal rahe  •  ${fmtSize(act.sumOf { it.speed })}/s  •  ${Prefs.connections} connections",
            fontSize = 12.sp, modifier = Modifier.fillMaxWidth().background(Color(0xFFDCE6F3)).padding(5.dp))
    }
}

@Composable
private fun TaskRow(t: Task, onPlay: (List<PItem>, Int) -> Unit) {
    val ctx = LocalContext.current
    var menu by remember { mutableStateOf(false) }
    var del by remember { mutableStateOf(false) }
    Box {
        Column(Modifier.fillMaxWidth().clickable { menu = true }.padding(horizontal = 8.dp, vertical = 7.dp)) {
            Row(verticalAlignment = Alignment.CenterVertically) {
                val (ic, tn) = fileIcon(File(t.name)); Icon(ic, null, tint = tn, modifier = Modifier.size(30.dp))
                Column(Modifier.weight(1f).padding(start = 8.dp)) {
                    Text(t.name, fontSize = 14.sp, maxLines = 1, overflow = TextOverflow.Ellipsis)
                    val status = when (t.state) {
                        DState.RUNNING -> (if (t.segTotal > 0) "Parts ${t.segDone}/${t.segTotal}" else "${fmtSize(t.done)} / ${if (t.total > 0) fmtSize(t.total) else "?"}") +
                            "  •  ${fmtSize(t.speed)}/s" + (if (t.speed > 0 && t.total > 0) "  •  ${eta((t.total - t.done) / t.speed)} left" else "") + "  •  ${t.conns} conn"
                        DState.QUEUED -> "Queue mai..."
                        DState.PAUSED -> "Paused  •  ${fmtSize(t.done)}" + (if (t.total > 0) " / ${fmtSize(t.total)}" else "")
                        DState.DONE -> "Complete  •  ${fmtSize(t.file.length())}"
                        DState.FAILED -> "Fail: ${t.error}"
                        DState.CONVERTING -> "MP4 bana raha hai... ${t.segDone}%"
                    }
                    Text(status, fontSize = 11.sp, color = if (t.state == DState.FAILED) Color(0xFFC13A1F) else W7.gray, maxLines = 2)
                }
                when (t.state) {
                    DState.RUNNING, DState.QUEUED -> IconButton({ Dl.pause(t) }) { Icon(Icons.Filled.Pause, "Pause") }
                    DState.PAUSED, DState.FAILED -> IconButton({ Dl.resume(t) }) { Icon(Icons.Filled.PlayArrow, "Resume", tint = Color(0xFF2E7D32)) }
                    DState.DONE -> IconButton({ openTask(ctx, t, onPlay) }) { Icon(Icons.Filled.OpenInNew, "Open", tint = W7.link) }
                    else -> {}
                }
            }
            if (t.state != DState.DONE) WinProgress(t.fraction, Modifier.padding(top = 4.dp), 12)
        }
        DropdownMenu(menu, { menu = false }) {
            if (t.state == DState.DONE) {
                DropdownMenuItem({ Text("Open") }, { menu = false; openTask(ctx, t, onPlay) })
                DropdownMenuItem({ Text("Open folder") }, { menu = false; ExplorerNav.request = Place.Dir(File(t.dir)); UiState.show(Screen.EXPLORER) })
                DropdownMenuItem({ Text("Share") }, { menu = false; Fs.share(ctx, listOf(t.file)) })
            } else if (Fs.isMedia(File(t.name)) && !t.isHls) {
                DropdownMenuItem({ Text("Play while downloading (stream)") }, { menu = false; onPlay(listOf(PItem(Uri.parse(t.url), t.name, Fs.kind(File(t.name)) == Kind.AUDIO)), 0) })
            }
            DropdownMenuItem({ Text("Restart") }, { menu = false; Dl.restart(t) })
            DropdownMenuItem({ Text("Copy link") }, {
                menu = false
                ctx.getSystemService(ClipboardManager::class.java).setPrimaryClip(android.content.ClipData.newPlainText("url", t.url))
            })
            DropdownMenuItem({ Text("Remove from list") }, { menu = false; Dl.remove(t, false) })
            DropdownMenuItem({ Text("Delete file bhi") }, { menu = false; del = true })
        }
    }
    if (del) WinDialog("Delete download", { del = false }, Icons.Filled.DeleteForever, Color(0xFFC13A1F), buttons = {
        WinButton("Yes", default = true) { del = false; Dl.remove(t, true) }
        WinButton("No") { del = false }
    }) { Text("'${t.name}' ki file bhi delete karni hai?", fontSize = 14.sp) }
}

private fun eta(s: Long): String = when {
    s < 60 -> "${s}s"; s < 3600 -> "${s / 60}m ${s % 60}s"; else -> "${s / 3600}h ${(s % 3600) / 60}m"
}

private fun openTask(ctx: Context, t: Task, onPlay: (List<PItem>, Int) -> Unit) {
    val f = t.file
    if (!f.exists()) return
    when {
        Fs.isMedia(f) -> onPlay(listOf(f.toPItem()), 0)
        Archive.supported(f) || Fs.kind(f) == Kind.IMAGE -> { f.parentFile?.let { ExplorerNav.request = Place.Dir(it) }; UiState.show(Screen.EXPLORER) }
        else -> Fs.openExternal(ctx, f)
    }
}

@Composable
fun SettingsDialog(onClose: () -> Unit) {
    WinDialog("Control Panel — FU XPlayer", { Prefs.save(); onClose() }, Icons.Filled.Settings, Color(0xFF607D8B), buttons = {
        WinButton("OK", default = true) { Prefs.save(); onClose() }
    }) {
        Text("Downloads", fontWeight = FontWeight.Bold, fontSize = 14.sp, color = Color(0xFF1E3A5F))
        Text("Connections per file: ${Prefs.connections}", fontSize = 13.sp)
        Slider(Prefs.connections.toFloat(), { Prefs.connections = it.toInt() }, valueRange = 1f..16f, steps = 14)
        Text("Ek saath downloads: ${Prefs.maxActive}", fontSize = 13.sp)
        Slider(Prefs.maxActive.toFloat(), { Prefs.maxActive = it.toInt() }, valueRange = 1f..6f, steps = 4)
        Toggle("HLS (.m3u8) ko MP4 mai badlo", Prefs.hlsToMp4) { Prefs.hlsToMp4 = it }
        Spacer(Modifier.height(6.dp))
        Text("Player & Browser", fontWeight = FontWeight.Bold, fontSize = 14.sp, color = Color(0xFF1E3A5F))
        Toggle("Background play (screen band ho to bhi chale)", Prefs.bgPlay) { Prefs.bgPlay = it }
        Toggle("Browser ad blocker", Prefs.adBlock) { Prefs.adBlock = it }
        Toggle("Explorer mai hidden files", Prefs.showHidden) { Prefs.showHidden = it }
        Spacer(Modifier.height(6.dp))
        Text("FU XPlayer v1.0 • Prince Faizan • Innovating Tomorrow", fontSize = 11.sp, color = W7.gray)
    }
}

@Composable
private fun Toggle(label: String, v: Boolean, onChange: (Boolean) -> Unit) {
    Row(Modifier.fillMaxWidth().clickable { onChange(!v) }, verticalAlignment = Alignment.CenterVertically) {
        Checkbox(v, onChange); Text(label, fontSize = 13.sp)
    }
}
