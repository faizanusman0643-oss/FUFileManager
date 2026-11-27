package com.fu.xplayer

import android.content.Context
import android.content.Intent
import android.net.Uri
import android.widget.Toast
import androidx.activity.compose.BackHandler
import androidx.compose.foundation.*
import androidx.compose.foundation.gestures.detectTapGestures
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.grid.GridCells
import androidx.compose.foundation.lazy.grid.LazyVerticalGrid
import androidx.compose.foundation.lazy.grid.items
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.pager.HorizontalPager
import androidx.compose.foundation.pager.rememberPagerState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Brush
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.graphicsLayer
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.layout.ContentScale
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.compose.ui.window.Dialog
import androidx.compose.ui.window.DialogProperties
import coil.compose.AsyncImage
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.io.File
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

sealed class Place {
    object Computer : Place()
    data class Dir(val f: File) : Place()
    data class Lib(val k: Kind) : Place()
    object Bin : Place()
    data class Search(val base: File, val q: String) : Place()
}

enum class SortBy { NAME, DATE, SIZE, TYPE }
enum class ViewMode { DETAILS, ICONS, TILES }

object Clip {
    var files by mutableStateOf<List<File>>(emptyList())
    var cut by mutableStateOf(false)
}

/** Progress for long operations (Win7 copy dialog). */
object Op {
    var running by mutableStateOf(false)
    var title by mutableStateOf("")
    var name by mutableStateOf("")
    var done by mutableLongStateOf(0L)
    var total by mutableLongStateOf(0L)
}

object ExplorerNav { var request by mutableStateOf<Place?>(null) }

private val dateFmt = SimpleDateFormat("dd/MM/yyyy h:mm a", Locale.US)

fun libTitle(k: Kind) = when (k) { Kind.VIDEO -> "Videos"; Kind.AUDIO -> "Music"; Kind.IMAGE -> "Pictures"; else -> "Documents" }

/** Opens a file the "right" way: media in our player, images in viewer, archives in archive viewer, else external app. */
class Opener(val ctx: Context, val play: (List<PItem>, Int) -> Unit, val viewImages: (List<File>, Int) -> Unit, val viewArchive: (File) -> Unit) {
    fun open(f: File, siblings: List<File> = listOf(f)) {
        when {
            Fs.isMedia(f) -> {
                val list = siblings.filter { Fs.isMedia(it) }.ifEmpty { listOf(f) }
                play(list.map { it.toPItem() }, list.indexOf(f).coerceAtLeast(0))
            }
            Fs.kind(f) == Kind.IMAGE -> { val l = siblings.filter { Fs.kind(it) == Kind.IMAGE }.ifEmpty { listOf(f) }; viewImages(l, l.indexOf(f).coerceAtLeast(0)) }
            Archive.supported(f) -> viewArchive(f)
            else -> Fs.openExternal(ctx, f)
        }
    }
}

@Composable
fun ExplorerScreen(onMin: () -> Unit, onClose: () -> Unit, onPlay: (List<PItem>, Int) -> Unit) {
    val ctx = LocalContext.current
    val scope = rememberCoroutineScope()
    var place by remember { mutableStateOf<Place>(Place.Computer) }
    val back = remember { mutableStateListOf<Place>() }
    val fwd = remember { mutableStateListOf<Place>() }
    var shown by remember { mutableStateOf<List<File>>(emptyList()) }
    var bin by remember { mutableStateOf<List<BinItem>>(emptyList()) }
    var loading by remember { mutableStateOf(false) }
    var reload by remember { mutableIntStateOf(0) }
    var sortBy by remember { mutableStateOf(SortBy.NAME) }
    var asc by remember { mutableStateOf(true) }
    var view by remember { mutableStateOf(ViewMode.DETAILS) }
    val sel = remember { mutableStateListOf<File>() }
    var navOpen by remember { mutableStateOf(false) }
    var searching by remember { mutableStateOf(false) }
    var query by remember { mutableStateOf("") }
    var menu by remember { mutableStateOf("") }
    var dialog by remember { mutableStateOf<String?>(null) }
    var target by remember { mutableStateOf<File?>(null) }
    var input by remember { mutableStateOf("") }
    var input2 by remember { mutableStateOf("") }
    var images by remember { mutableStateOf<Pair<List<File>, Int>?>(null) }
    var archive by remember { mutableStateOf<File?>(null) }
    var props by remember { mutableStateOf<List<File>?>(null) }
    var pwAsk by remember { mutableStateOf<Pair<File, Boolean>?>(null) }

    val opener = remember { Opener(ctx, onPlay, { l, i -> images = l to i }, { archive = it }) }

    fun go(p: Place, record: Boolean = true) {
        if (record && p != place) { back.add(place); fwd.clear() }
        place = p; sel.clear(); navOpen = false; searching = false
    }

    LaunchedEffect(ExplorerNav.request) { ExplorerNav.request?.let { go(it); ExplorerNav.request = null } }

    LaunchedEffect(place, reload, sortBy, asc, Prefs.showHidden) {
        loading = true
        val p = place
        val list = withContext(Dispatchers.IO) {
            when (p) {
                is Place.Dir -> Fs.list(p.f, Prefs.showHidden)
                is Place.Lib -> Fs.library(ctx, p.k)
                is Place.Search -> Fs.search(p.base, p.q, Prefs.showHidden)
                is Place.Bin -> { bin = Fs.binItems(); emptyList() }
                else -> emptyList()
            }
        }
        val c: Comparator<File> = when (sortBy) {
            SortBy.NAME -> compareBy { it.name.lowercase() }
            SortBy.DATE -> compareBy { it.lastModified() }
            SortBy.SIZE -> compareBy { if (it.isDirectory) 0L else it.length() }
            SortBy.TYPE -> compareBy({ it.extension.lowercase() }, { it.name.lowercase() })
        }
        val sorted = list.sortedWith(if (asc) c else c.reversed())
        shown = if (p is Place.Lib) sorted else sorted.sortedBy { !it.isDirectory }
        loading = false
    }

    BackHandler(enabled = sel.isNotEmpty() || back.isNotEmpty() || navOpen || searching) {
        when {
            navOpen -> navOpen = false
            searching -> searching = false
            sel.isNotEmpty() -> sel.clear()
            else -> { fwd.add(place); place = back.removeAt(back.lastIndex); sel.clear() }
        }
    }

    val curDir: File? = (place as? Place.Dir)?.f

    fun toast(s: String) = Toast.makeText(ctx, s, Toast.LENGTH_SHORT).show()

    fun runOp(title: String, notify: Boolean = false, work: suspend ((Long, Long, String) -> Unit) -> String) {
        Op.title = title; Op.name = ""; Op.done = 0; Op.total = 0; Op.running = true
        val t0 = System.currentTimeMillis()
        scope.launch {
            val msg = withContext(Dispatchers.IO) {
                try { work { d, t, n -> Op.done = d; Op.total = t; Op.name = n } }
                catch (e: NeedPassword) { "!pw" }
                catch (e: Exception) { "Error: ${e.message}" }
            }
            Op.running = false
            if (msg != "!pw") {
                toast(msg)
                if (notify || System.currentTimeMillis() - t0 > 4000) Notify.show(ctx, 9000 + (t0 % 999).toInt(), "✔ $title complete", msg, null)
            }
            sel.clear(); reload++
        }
    }

    fun paste(dest: File) {
        val files = Clip.files; val cut = Clip.cut
        if (files.isEmpty()) return
        runOp(if (cut) "Moving ${files.size} items" else "Copying ${files.size} items") { p ->
            val r = Fs.transfer(files, dest, cut, p)
            Fs.scan(ctx, r.map { it.absolutePath } + files.map { it.absolutePath })
            if (cut) withContext(Dispatchers.Main) { Clip.files = emptyList() }
            "${r.size} items ${if (cut) "move" else "copy"} ho gaye"
        }
    }

    fun extract(f: File, pw: String?) {
        val dest = f.parentFile ?: Fs.root
        runOp("Extracting ${f.name}", notify = true) { p ->
            try {
                val out = Archive.extract(f, dest, null, pw) { n -> p(0, 0, n) }
                Fs.scan(ctx, listOf(out.absolutePath))
                "Extract ho gaya: ${out.name}"
            } catch (e: NeedPassword) {
                withContext(Dispatchers.Main) { pwAsk = f to e.wrong }; throw e
            }
        }
    }

    fun delete(files: List<File>, permanent: Boolean) = runOp("Deleting ${files.size} items") {
        if (permanent) { files.forEach { Fs.deleteRec(it) } } else Fs.toBin(files)
        Fs.scan(ctx, files.map { it.absolutePath })
        if (permanent) "${files.size} permanently delete" else "${files.size} Recycle Bin mai chale gaye"
    }

    WinWindow(
        title = when (val p = place) {
            Place.Computer -> "Computer"; is Place.Dir -> p.f.name.ifEmpty { "Internal storage" }
            is Place.Lib -> libTitle(p.k); Place.Bin -> "Recycle Bin"; is Place.Search -> "Search Results"
        },
        icon = if (place == Place.Bin) Icons.Filled.Delete else Icons.Filled.FolderOpen, iconTint = W7.folder,
        onMin = onMin, onClose = onClose
    ) {
        // ---- address bar ----
        Row(Modifier.fillMaxWidth().background(W7.aero).padding(3.dp), verticalAlignment = Alignment.CenterVertically) {
            RoundNav(Icons.Filled.ArrowBack, back.isNotEmpty()) { fwd.add(place); place = back.removeAt(back.lastIndex); sel.clear() }
            RoundNav(Icons.Filled.ArrowForward, fwd.isNotEmpty()) { back.add(place); place = fwd.removeAt(fwd.lastIndex); sel.clear() }
            RoundNav(Icons.Filled.ArrowUpward, curDir != null) {
                val par = curDir?.parentFile
                if (par != null && curDir?.absolutePath != Fs.root.absolutePath && par.canRead()) go(Place.Dir(par)) else go(Place.Computer)
            }
            if (searching) {
                WinField(query, { query = it }, Modifier.weight(1f), "Search ${curDir?.name ?: "Computer"}", onDone = {
                    if (query.isNotBlank()) go(Place.Search(curDir ?: Fs.root, query))
                })
                IconButton({ searching = false }, Modifier.size(34.dp)) { Icon(Icons.Filled.Close, "Close search") }
            } else {
                Breadcrumb(place, Modifier.weight(1f)) { go(it) }
                IconButton({ searching = true; query = "" }, Modifier.size(34.dp)) { Icon(Icons.Filled.Search, "Search", tint = W7.link) }
            }
        }
        // ---- command bar ----
        Row(
            Modifier.fillMaxWidth().background(W7.cmd).horizontalScroll(rememberScrollState()).padding(horizontal = 2.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            CmdItem(null, Icons.Filled.Menu) { navOpen = !navOpen }
            if (sel.isNotEmpty()) {
                Text("${sel.size} selected", fontSize = 13.sp, color = W7.link, modifier = Modifier.padding(horizontal = 6.dp))
                CmdItem("Cut", Icons.Filled.ContentCut) { Clip.files = sel.toList(); Clip.cut = true; sel.clear(); toast("Cut — ab folder khol ke Paste karen") }
                CmdItem("Copy", Icons.Filled.ContentCopy) { Clip.files = sel.toList(); Clip.cut = false; sel.clear(); toast("Copied — ab folder khol ke Paste karen") }
                CmdItem("Delete", Icons.Filled.Delete, Color(0xFFC13A1F)) { dialog = "delete" }
                if (sel.size == 1) CmdItem("Rename", Icons.Filled.Edit) { target = sel[0]; input = sel[0].name; dialog = "rename" }
                CmdItem("Share", Icons.Filled.Share) { Fs.share(ctx, sel.toList()) }
                Box {
                    CmdItem("More ▾", null) { menu = "more" }
                    DropdownMenu(menu == "more", { menu = "" }) {
                        DropdownMenuItem({ Text("Select all") }, { menu = ""; sel.clear(); sel.addAll(shown) }, leadingIcon = { Icon(Icons.Filled.SelectAll, null) })
                        DropdownMenuItem({ Text("Add to ZIP") }, { menu = ""; input = (sel.singleOrNull()?.nameWithoutExtension ?: "Archive") + ".zip"; input2 = ""; dialog = "zip" }, leadingIcon = { Icon(Icons.Filled.Archive, null) })
                        DropdownMenuItem({ Text("Add to ZIP with password") }, { menu = ""; input = (sel.singleOrNull()?.nameWithoutExtension ?: "Archive") + ".zip"; input2 = ""; dialog = "zippw" }, leadingIcon = { Icon(Icons.Filled.Lock, null) })
                        if (sel.size == 1 && Archive.supported(sel[0])) DropdownMenuItem({ Text("Extract here") }, { menu = ""; extract(sel[0], null) }, leadingIcon = { Icon(Icons.Filled.Unarchive, null) })
                        if (sel.all { Fs.isMedia(it) }) DropdownMenuItem({ Text("Play all") }, { menu = ""; onPlay(sel.map { it.toPItem() }, 0); sel.clear() }, leadingIcon = { Icon(Icons.Filled.PlayArrow, null) })
                        if (sel.size == 1 && sel[0].isFile) DropdownMenuItem({ Text("Open with...") }, { menu = ""; Fs.openExternal(ctx, sel[0], true) }, leadingIcon = { Icon(Icons.Filled.OpenInNew, null) })
                        DropdownMenuItem({ Text("Delete permanently") }, { menu = ""; dialog = "delete!" }, leadingIcon = { Icon(Icons.Filled.DeleteForever, null) })
                        DropdownMenuItem({ Text("Properties") }, { menu = ""; props = sel.toList() }, leadingIcon = { Icon(Icons.Filled.Info, null) })
                    }
                }
            } else if (place == Place.Bin) {
                CmdItem("Empty Recycle Bin", Icons.Filled.DeleteSweep, Color(0xFFC13A1F)) { dialog = "emptybin" }
                CmdItem("Restore all items", Icons.Filled.Restore) { runOp("Restoring") { "${Fs.restore(Fs.binItems())} items restore ho gaye" } }
            } else {
                Box {
                    CmdItem("Organize ▾", null) { menu = "org" }
                    DropdownMenu(menu == "org", { menu = "" }) {
                        DropdownMenuItem({ Text("Paste") }, { menu = ""; curDir?.let { paste(it) } }, enabled = Clip.files.isNotEmpty() && curDir != null, leadingIcon = { Icon(Icons.Filled.ContentPaste, null) })
                        DropdownMenuItem({ Text("Select all") }, { menu = ""; sel.addAll(shown) }, leadingIcon = { Icon(Icons.Filled.SelectAll, null) })
                        DropdownMenuItem({ Text(if (Prefs.showHidden) "✓ Hidden files dikhao" else "Hidden files dikhao") }, { menu = ""; Prefs.showHidden = !Prefs.showHidden; Prefs.save() }, leadingIcon = { Icon(Icons.Filled.Visibility, null) })
                        DropdownMenuItem({ Text("Refresh") }, { menu = ""; reload++ }, leadingIcon = { Icon(Icons.Filled.Refresh, null) })
                        if (curDir != null) DropdownMenuItem({ Text("Properties") }, { menu = ""; props = listOf(curDir) }, leadingIcon = { Icon(Icons.Filled.Info, null) })
                    }
                }
                if (curDir != null) CmdItem("New folder", Icons.Filled.CreateNewFolder, W7.folder) { input = Fs.unique(curDir, "New folder").name; dialog = "newfolder" }
                if (Clip.files.isNotEmpty() && curDir != null) CmdItem("Paste (${Clip.files.size})", Icons.Filled.ContentPaste) { paste(curDir) }
                Box {
                    CmdItem("View ▾", if (view == ViewMode.DETAILS) Icons.Filled.ViewList else Icons.Filled.GridView) { menu = "view" }
                    DropdownMenu(menu == "view", { menu = "" }) {
                        DropdownMenuItem({ Text("Details") }, { menu = ""; view = ViewMode.DETAILS })
                        DropdownMenuItem({ Text("Large icons") }, { menu = ""; view = ViewMode.ICONS })
                        DropdownMenuItem({ Text("Tiles") }, { menu = ""; view = ViewMode.TILES })
                    }
                }
                Box {
                    CmdItem("Sort ▾", Icons.Filled.Sort) { menu = "sort" }
                    DropdownMenu(menu == "sort", { menu = "" }) {
                        SortBy.entries.forEach { s ->
                            DropdownMenuItem({ Text((if (s == sortBy) "● " else "   ") + s.name.lowercase().replaceFirstChar { it.uppercase() }) }, { menu = ""; sortBy = s })
                        }
                        HorizontalDivider()
                        DropdownMenuItem({ Text(if (asc) "● Ascending" else "   Ascending") }, { menu = ""; asc = true })
                        DropdownMenuItem({ Text(if (!asc) "● Descending" else "   Descending") }, { menu = ""; asc = false })
                    }
                }
                if (shown.any { Fs.isMedia(it) }) CmdItem("Play all", Icons.Filled.PlayArrow, Color(0xFF2E7D32)) {
                    val m = shown.filter { Fs.isMedia(it) }; onPlay(m.map { it.toPItem() }, 0)
                }
            }
        }
        HorizontalDivider(color = Color(0xFFC5D2E0))
        // ---- body ----
        Box(Modifier.weight(1f).fillMaxWidth()) {
            when (place) {
                Place.Computer -> ComputerView(ctx) { go(it) }
                Place.Bin -> BinView(bin, onRestore = { b -> runOp("Restoring") { "${Fs.restore(listOf(b))} restore hua" } },
                    onPurge = { b -> runOp("Deleting") { Fs.purge(listOf(b)); "Permanently delete" } })
                else -> {
                    if (!loading && shown.isEmpty()) {
                        Text(if (place is Place.Search) "Koi item nahi mila." else "This folder is empty.", color = W7.gray,
                            modifier = Modifier.align(Alignment.TopCenter).padding(24.dp))
                    }
                    val click: (File) -> Unit = { f ->
                        if (sel.isNotEmpty()) { if (!sel.remove(f)) sel.add(f) }
                        else if (f.isDirectory) go(Place.Dir(f)) else opener.open(f, shown)
                    }
                    val toggle: (File) -> Unit = { f -> if (!sel.remove(f)) sel.add(f) }
                    val ctxMenu: @Composable (File, Boolean, () -> Unit) -> Unit = { f, show, dismiss ->
                        DropdownMenu(show, dismiss) {
                            DropdownMenuItem({ Text("Open", fontWeight = FontWeight.Bold) }, { dismiss(); if (f.isDirectory) go(Place.Dir(f)) else opener.open(f, shown) })
                            if (f.isFile) DropdownMenuItem({ Text("Open with...") }, { dismiss(); Fs.openExternal(ctx, f, true) })
                            if (Archive.supported(f)) {
                                DropdownMenuItem({ Text("Open archive") }, { dismiss(); archive = f })
                                DropdownMenuItem({ Text("Extract here") }, { dismiss(); extract(f, null) })
                            }
                            if (place is Place.Search || place is Place.Lib) DropdownMenuItem({ Text("Open file location") }, { dismiss(); f.parentFile?.let { go(Place.Dir(it)) } })
                            HorizontalDivider()
                            DropdownMenuItem({ Text("Select") }, { dismiss(); toggle(f) })
                            DropdownMenuItem({ Text("Cut") }, { dismiss(); Clip.files = listOf(f); Clip.cut = true })
                            DropdownMenuItem({ Text("Copy") }, { dismiss(); Clip.files = listOf(f); Clip.cut = false })
                            if (f.isDirectory && Clip.files.isNotEmpty()) DropdownMenuItem({ Text("Paste into") }, { dismiss(); paste(f) })
                            HorizontalDivider()
                            DropdownMenuItem({ Text("Add to ZIP") }, { dismiss(); sel.clear(); sel.add(f); input = f.nameWithoutExtension + ".zip"; input2 = ""; dialog = "zip" })
                            DropdownMenuItem({ Text("Add to ZIP with password") }, { dismiss(); sel.clear(); sel.add(f); input = f.nameWithoutExtension + ".zip"; input2 = ""; dialog = "zippw" })
                            DropdownMenuItem({ Text("Share") }, { dismiss(); Fs.share(ctx, listOf(f)) })
                            DropdownMenuItem({ Text("Delete") }, { dismiss(); sel.clear(); sel.add(f); dialog = "delete" })
                            DropdownMenuItem({ Text("Rename") }, { dismiss(); target = f; input = f.name; dialog = "rename" })
                            HorizontalDivider()
                            DropdownMenuItem({ Text("Properties") }, { dismiss(); props = listOf(f) })
                        }
                    }
                    when (view) {
                        ViewMode.DETAILS -> LazyColumn(Modifier.fillMaxSize()) {
                            items(shown, key = { it.absolutePath }) { f ->
                                FileRow(f, f in sel, place is Place.Search || place is Place.Lib, { click(f) }, { toggle(f) }, ctxMenu)
                            }
                        }
                        else -> LazyVerticalGrid(GridCells.Adaptive(if (view == ViewMode.ICONS) 104.dp else 160.dp), Modifier.fillMaxSize()) {
                            items(shown, key = { it.absolutePath }) { f -> FileTile(f, f in sel, view == ViewMode.TILES, { click(f) }, { toggle(f) }, ctxMenu) }
                        }
                    }
                }
            }
            if (loading) LinearProgressIndicator(Modifier.fillMaxWidth().align(Alignment.TopCenter), color = W7.teal)
            if (navOpen) {
                Row(Modifier.fillMaxSize()) {
                    NavPane(ctx, Modifier.width(230.dp).fillMaxHeight()) { go(it) }
                    Box(Modifier.weight(1f).fillMaxHeight().background(Color(0x33000000)).clickable { navOpen = false })
                }
            }
        }
        // ---- status / details bar ----
        Row(
            Modifier.fillMaxWidth().background(Brush.verticalGradient(listOf(Color(0xFFF1F6FB), Color(0xFFDCE6F3)))).padding(horizontal = 8.dp, vertical = 5.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            val single = sel.singleOrNull()
            if (single != null) {
                val (ic, tn) = fileIcon(single)
                Icon(ic, null, tint = tn, modifier = Modifier.size(28.dp))
                Column(Modifier.padding(start = 8.dp)) {
                    Text(single.name, fontSize = 12.sp, maxLines = 1, overflow = TextOverflow.Ellipsis)
                    Text("${Fs.typeName(single)}  •  ${if (single.isFile) fmtSize(single.length()) else ""}  •  ${dateFmt.format(Date(single.lastModified()))}", fontSize = 11.sp, color = W7.gray, maxLines = 1)
                }
            } else {
                val n = if (place == Place.Bin) bin.size else shown.size
                Text(if (sel.isNotEmpty()) "${sel.size} items selected  •  ${fmtSize(sel.filter { it.isFile }.sumOf { it.length() })}" else "$n items", fontSize = 12.sp, color = W7.text)
                if (Clip.files.isNotEmpty()) Text("   📋 ${Clip.files.size} ${if (Clip.cut) "cut" else "copied"}", fontSize = 12.sp, color = W7.link)
            }
        }
    }

    // ---------------- dialogs ----------------
    when (dialog) {
        "newfolder", "rename" -> WinDialog(if (dialog == "rename") "Rename" else "New folder", { dialog = null }, Icons.Filled.Edit, W7.folder, buttons = {
            WinButton("OK", default = true) {
                val name = Dl.clean(input.trim())
                if (dialog == "rename") {
                    val t = target
                    if (t != null && name != t.name) {
                        val dest = File(t.parentFile, name)
                        if (dest.exists()) toast("Is naam ki file pehle se hai") else if (t.renameTo(dest)) { Fs.scan(ctx, listOf(t.absolutePath, dest.absolutePath)); sel.clear() } else toast("Rename fail")
                    }
                } else curDir?.let { if (!Fs.unique(it, name).mkdirs()) toast("Folder nahi bana") }
                dialog = null; reload++
            }
            WinButton("Cancel") { dialog = null }
        }) {
            Text(if (dialog == "rename") "Naya naam:" else "Folder ka naam:", fontSize = 14.sp)
            Spacer(Modifier.height(6.dp)); WinField(input, { input = it })
        }
        "delete", "delete!" -> {
            val perm = dialog == "delete!" || sel.any { it.absolutePath.startsWith(Fs.bin.absolutePath) }
            WinDialog("Delete ${if (sel.size == 1) "File" else "Multiple Items"}", { dialog = null }, if (perm) Icons.Filled.DeleteForever else Icons.Filled.Delete, Color(0xFFC13A1F), buttons = {
                WinButton("Yes", default = true) { val l = sel.toList(); dialog = null; delete(l, perm) }
                if (!perm) WinButton("Permanently") { val l = sel.toList(); dialog = null; delete(l, true) }
                WinButton("No") { dialog = null }
            }) {
                Text(
                    if (perm) "Kya aap ${if (sel.size == 1) "'${sel[0].name}'" else "in ${sel.size} items"} ko permanently delete karna chahte hain?"
                    else "Kya aap ${if (sel.size == 1) "'${sel[0].name}'" else "in ${sel.size} items"} ko Recycle Bin mai bhejna chahte hain?",
                    fontSize = 14.sp
                )
            }
        }
        "emptybin" -> WinDialog("Delete Multiple Items", { dialog = null }, Icons.Filled.DeleteSweep, Color(0xFFC13A1F), buttons = {
            WinButton("Yes", default = true) { dialog = null; runOp("Emptying Recycle Bin") { Fs.emptyBin(); "Recycle Bin khali ho gaya" } }
            WinButton("No") { dialog = null }
        }) { Text("Kya aap Recycle Bin ke ${bin.size} items permanently delete karna chahte hain?", fontSize = 14.sp) }
        "zip", "zippw" -> WinDialog(if (dialog == "zippw") "Add to ZIP (password)" else "Add to ZIP", { dialog = null }, Icons.Filled.Archive, Color(0xFFB8860B), buttons = {
            WinButton("OK", default = true) {
                val files = sel.toList(); val pw = if (dialog == "zippw") input2 else null
                if (dialog == "zippw" && input2.length < 3) { toast("Password kam se kam 3 characters"); return@WinButton }
                val parent = files.firstOrNull()?.parentFile ?: curDir ?: Fs.root
                val out = Fs.unique(parent, Dl.clean(input.trim()).let { if (it.endsWith(".zip", true)) it else "$it.zip" })
                dialog = null
                runOp("Compressing ${files.size} items", notify = true) { p ->
                    Archive.createZip(files, out, pw) { n -> p(0, 0, n) }
                    Fs.scan(ctx, listOf(out.absolutePath)); "ZIP ban gayi: ${out.name}${if (pw != null) " (password ke saath)" else ""}"
                }
            }
            WinButton("Cancel") { dialog = null }
        }) {
            Text("Archive ka naam:", fontSize = 14.sp); Spacer(Modifier.height(4.dp)); WinField(input, { input = it })
            if (dialog == "zippw") {
                Spacer(Modifier.height(8.dp)); Text("Password (AES-256):", fontSize = 14.sp); Spacer(Modifier.height(4.dp))
                WinField(input2, { input2 = it }, hint = "Password likhen", password = true)
                Text("Note: RAR banana WinRAR ka proprietary format hai, isliye password ZIP banti hai — WinRAR/7-Zip dono isay khol lete hain.", fontSize = 11.sp, color = W7.gray)
            }
        }
    }

    pwAsk?.let { (f, wrong) ->
        PasswordDialog(f.name, wrong, { pwAsk = null }) { pw -> pwAsk = null; extract(f, pw) }
    }

    props?.let { PropertiesDialog(it) { props = null } }
    images?.let { (l, i) -> ImageViewer(l, i) { images = null } }
    archive?.let { a -> ArchiveViewer(a, opener, { archive = null; reload++ }) }
    if (Op.running) OpDialog()
}

@Composable
private fun RoundNav(icon: androidx.compose.ui.graphics.vector.ImageVector, enabled: Boolean, onClick: () -> Unit) {
    Box(
        Modifier.padding(1.dp).size(32.dp).clip(RoundedCornerShape(16.dp))
            .background(if (enabled) Brush.verticalGradient(listOf(Color(0xFF5FA8E8), Color(0xFF1C5FA8))) else Brush.verticalGradient(listOf(Color(0xFFD0DAE6), Color(0xFFB0BFD0))))
            .clickable(enabled = enabled, onClick = onClick),
        contentAlignment = Alignment.Center
    ) { Icon(icon, null, tint = Color.White, modifier = Modifier.size(18.dp)) }
}

@Composable
private fun Breadcrumb(place: Place, modifier: Modifier, onGo: (Place) -> Unit) {
    val parts = ArrayList<Pair<String, Place>>()
    parts += "Computer" to Place.Computer
    when (place) {
        is Place.Dir -> {
            val root = Fs.root.absolutePath
            val p = place.f.absolutePath
            if (p.startsWith(root)) {
                parts += "Internal storage" to Place.Dir(Fs.root)
                var cur = Fs.root
                p.removePrefix(root).split('/').filter { it.isNotEmpty() }.forEach { s -> cur = File(cur, s); parts += s to Place.Dir(cur) }
            } else {
                var cur = File("/")
                p.split('/').filter { it.isNotEmpty() }.forEach { s -> cur = File(cur, s); parts += s to Place.Dir(cur) }
            }
        }
        is Place.Lib -> { parts.clear(); parts += "Libraries" to Place.Computer; parts += libTitle(place.k) to place }
        Place.Bin -> { parts.clear(); parts += "Recycle Bin" to Place.Bin }
        is Place.Search -> parts += "Search: \"${place.q}\"" to place
        else -> {}
    }
    val sc = rememberScrollState()
    LaunchedEffect(parts.size) { sc.scrollTo(sc.maxValue) }
    Row(
        modifier.height(34.dp).background(Color.White).border(1.dp, Color(0xFF8E9FB4)).horizontalScroll(sc).padding(horizontal = 4.dp),
        verticalAlignment = Alignment.CenterVertically
    ) {
        parts.forEachIndexed { i, (label, p) ->
            if (i > 0) Text(" ▸ ", fontSize = 12.sp, color = W7.gray)
            Text(label, fontSize = 13.sp, color = W7.text, modifier = Modifier.clip(RoundedCornerShape(2.dp)).clickable { onGo(p) }.padding(horizontal = 3.dp, vertical = 6.dp))
        }
    }
}

@Composable
private fun NavPane(ctx: Context, modifier: Modifier, onGo: (Place) -> Unit) {
    val drives = remember { Fs.drives(ctx) }
    Column(modifier.background(W7.navBg).border(1.dp, Color(0xFFC5D2E0)).verticalScroll(rememberScrollState()).padding(vertical = 4.dp)) {
        @Composable fun head(t: String, icon: androidx.compose.ui.graphics.vector.ImageVector, tint: Color, p: Place) =
            NavLine(t, icon, tint, 0, true) { onGo(p) }
        head("Favorites", Icons.Filled.Star, Color(0xFFE8B931), Place.Dir(Fs.root))
        NavLine("Downloads", Icons.Filled.Download, W7.link, 1) { onGo(Place.Dir(Fs.downloads)) }
        NavLine("FU XPlayer Downloads", Icons.Filled.Download, W7.teal, 1) { Dl.base.mkdirs(); onGo(Place.Dir(Dl.base)) }
        NavLine("DCIM / Camera", Icons.Filled.PhotoCamera, Color(0xFF2E9E47), 1) { onGo(Place.Dir(File(Fs.root, "DCIM"))) }
        NavLine("WhatsApp Media", Icons.Filled.Chat, Color(0xFF25D366), 1) {
            val a = File(Fs.root, "Android/media/com.whatsapp/WhatsApp/Media"); val b = File(Fs.root, "WhatsApp/Media")
            onGo(Place.Dir(if (a.exists()) a else b))
        }
        Spacer(Modifier.height(8.dp))
        head("Libraries", Icons.Filled.LibraryBooks, Color(0xFFE8B931), Place.Computer)
        NavLine("Documents", Icons.Filled.Description, Color(0xFF2B66C4), 1) { onGo(Place.Lib(Kind.DOC)) }
        NavLine("Music", Icons.Filled.MusicNote, Color(0xFFE0533D), 1) { onGo(Place.Lib(Kind.AUDIO)) }
        NavLine("Pictures", Icons.Filled.Image, Color(0xFF2E9E47), 1) { onGo(Place.Lib(Kind.IMAGE)) }
        NavLine("Videos", Icons.Filled.Movie, Color(0xFF6A3FB5), 1) { onGo(Place.Lib(Kind.VIDEO)) }
        Spacer(Modifier.height(8.dp))
        head("Computer", Icons.Filled.Computer, W7.link, Place.Computer)
        drives.forEach { d -> NavLine(d.label.substringBefore(" —"), if (d.root == Fs.root) Icons.Filled.Storage else Icons.Filled.SdCard, Color(0xFF607D8B), 1) { onGo(Place.Dir(d.root)) } }
        Spacer(Modifier.height(8.dp))
        head("Recycle Bin", Icons.Filled.Delete, Color(0xFF607D8B), Place.Bin)
    }
}

@Composable
private fun NavLine(t: String, icon: androidx.compose.ui.graphics.vector.ImageVector, tint: Color, level: Int, bold: Boolean = false, onClick: () -> Unit) {
    Row(Modifier.fillMaxWidth().clickable(onClick = onClick).padding(start = (8 + level * 18).dp, top = 7.dp, bottom = 7.dp, end = 6.dp), verticalAlignment = Alignment.CenterVertically) {
        Icon(icon, null, tint = tint, modifier = Modifier.size(20.dp))
        Text(t, Modifier.padding(start = 8.dp), fontSize = 13.sp, color = if (bold) Color(0xFF1E3A5F) else W7.text, fontWeight = if (bold) FontWeight.SemiBold else FontWeight.Normal, maxLines = 1)
    }
}

@Composable
private fun ComputerView(ctx: Context, onGo: (Place) -> Unit) {
    val drives = remember { Fs.drives(ctx) }
    Column(Modifier.fillMaxSize().verticalScroll(rememberScrollState()).padding(10.dp)) {
        Text("Hard Disk Drives (${drives.size})", color = Color(0xFF1E3A5F), fontSize = 14.sp)
        HorizontalDivider(Modifier.padding(vertical = 4.dp), color = Color(0xFFC5D2E0))
        drives.forEach { d ->
            Row(Modifier.fillMaxWidth().clip(RoundedCornerShape(3.dp)).clickable { onGo(Place.Dir(d.root)) }.padding(8.dp), verticalAlignment = Alignment.CenterVertically) {
                Icon(if (d.root == Fs.root) Icons.Filled.Storage else Icons.Filled.SdCard, null, tint = Color(0xFF607D8B), modifier = Modifier.size(44.dp))
                Column(Modifier.padding(start = 10.dp).weight(1f)) {
                    Text(d.label, fontSize = 14.sp)
                    val used = if (d.total > 0) 1f - d.free.toFloat() / d.total else 0f
                    Box(Modifier.padding(vertical = 4.dp).fillMaxWidth().height(14.dp).background(Color(0xFFE6E6E6)).border(1.dp, Color(0xFFBCBCBC))) {
                        Box(Modifier.fillMaxHeight().fillMaxWidth(used).padding(1.dp).background(
                            if (used > 0.9f) Brush.verticalGradient(listOf(Color(0xFFF08080), Color(0xFFD32F2F)))
                            else Brush.verticalGradient(listOf(Color(0xFF7FC4F5), Color(0xFF2A8BD0)))
                        ))
                    }
                    Text("${fmtSize(d.free)} free of ${fmtSize(d.total)}", fontSize = 12.sp, color = W7.gray)
                }
            }
        }
        Spacer(Modifier.height(14.dp))
        Text("Libraries", color = Color(0xFF1E3A5F), fontSize = 14.sp)
        HorizontalDivider(Modifier.padding(vertical = 4.dp), color = Color(0xFFC5D2E0))
        Row(Modifier.horizontalScroll(rememberScrollState())) {
            listOf(Kind.DOC, Kind.AUDIO, Kind.IMAGE, Kind.VIDEO).forEach { k ->
                val (ic, tn) = when (k) { Kind.DOC -> Icons.Filled.Description to Color(0xFF2B66C4); Kind.AUDIO -> Icons.Filled.MusicNote to Color(0xFFE0533D); Kind.IMAGE -> Icons.Filled.Image to Color(0xFF2E9E47); else -> Icons.Filled.Movie to Color(0xFF6A3FB5) }
                Column(Modifier.width(86.dp).clip(RoundedCornerShape(3.dp)).clickable { onGo(Place.Lib(k)) }.padding(6.dp), horizontalAlignment = Alignment.CenterHorizontally) {
                    Icon(ic, null, tint = tn, modifier = Modifier.size(40.dp)); Text(libTitle(k), fontSize = 12.sp)
                }
            }
            Column(Modifier.width(86.dp).clip(RoundedCornerShape(3.dp)).clickable { onGo(Place.Bin) }.padding(6.dp), horizontalAlignment = Alignment.CenterHorizontally) {
                Icon(Icons.Filled.Delete, null, tint = Color(0xFF607D8B), modifier = Modifier.size(40.dp)); Text("Recycle Bin", fontSize = 12.sp)
            }
        }
    }
}

@Composable
private fun Thumb(f: File, size: Int) {
    val k = Fs.kind(f)
    if (k == Kind.IMAGE || k == Kind.VIDEO) {
        Box(Modifier.size(size.dp), contentAlignment = Alignment.Center) {
            val (ic, tn) = fileIcon(f)
            Icon(ic, null, tint = tn, modifier = Modifier.size((size * 0.8f).dp))
            AsyncImage(f, null, Modifier.size(size.dp).clip(RoundedCornerShape(2.dp)), contentScale = ContentScale.Crop)
            if (k == Kind.VIDEO) Icon(Icons.Filled.PlayCircle, null, tint = Color(0xCCFFFFFF), modifier = Modifier.size((size * 0.4f).dp))
        }
    } else {
        val (ic, tn) = fileIcon(f)
        Icon(ic, null, tint = tn, modifier = Modifier.size(size.dp))
    }
}

@OptIn(ExperimentalFoundationApi::class)
@Composable
private fun FileRow(f: File, selected: Boolean, showPath: Boolean, onClick: () -> Unit, onToggle: () -> Unit, ctxMenu: @Composable (File, Boolean, () -> Unit) -> Unit) {
    var menu by remember { mutableStateOf(false) }
    Box {
        Row(
            Modifier.fillMaxWidth().background(if (selected) W7.sel else Color.White)
                .border(if (selected) 1.dp else 0.dp, if (selected) W7.selBorder else Color.Transparent)
                .combinedClickable(onClick = onClick, onLongClick = { menu = true })
                .padding(horizontal = 6.dp, vertical = 5.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            Box(Modifier.clickable(onClick = onToggle)) {
                Thumb(f, 36)
                if (selected) Icon(Icons.Filled.CheckBox, null, tint = W7.link, modifier = Modifier.size(16.dp).align(Alignment.TopStart).background(Color.White))
            }
            Column(Modifier.weight(1f).padding(start = 8.dp)) {
                Text(f.name, fontSize = 14.sp, maxLines = 1, overflow = TextOverflow.Ellipsis, color = W7.text)
                Text(
                    if (showPath) f.parent ?: "" else "${dateFmt.format(Date(f.lastModified()))}   ${Fs.typeName(f)}",
                    fontSize = 11.sp, color = W7.gray, maxLines = 1, overflow = TextOverflow.Ellipsis
                )
            }
            Text(if (f.isFile) fmtSize(f.length()) else "", fontSize = 12.sp, color = W7.gray)
        }
        ctxMenu(f, menu) { menu = false }
    }
}

@OptIn(ExperimentalFoundationApi::class)
@Composable
private fun FileTile(f: File, selected: Boolean, tiles: Boolean, onClick: () -> Unit, onToggle: () -> Unit, ctxMenu: @Composable (File, Boolean, () -> Unit) -> Unit) {
    var menu by remember { mutableStateOf(false) }
    Box(Modifier.padding(3.dp)) {
        val m = Modifier.fillMaxWidth().clip(RoundedCornerShape(3.dp)).background(if (selected) W7.sel else Color.Transparent)
            .border(1.dp, if (selected) W7.selBorder else Color.Transparent, RoundedCornerShape(3.dp))
            .combinedClickable(onClick = onClick, onLongClick = { menu = true }).padding(6.dp)
        if (tiles) Row(m, verticalAlignment = Alignment.CenterVertically) {
            Box(Modifier.clickable(onClick = onToggle)) { Thumb(f, 48) }
            Column(Modifier.padding(start = 6.dp)) {
                Text(f.name, fontSize = 13.sp, maxLines = 2, overflow = TextOverflow.Ellipsis)
                Text(Fs.typeName(f), fontSize = 11.sp, color = W7.gray, maxLines = 1)
                if (f.isFile) Text(fmtSize(f.length()), fontSize = 11.sp, color = W7.gray)
            }
        } else Column(m, horizontalAlignment = Alignment.CenterHorizontally) {
            Box(Modifier.clickable(onClick = onToggle)) { Thumb(f, 72) }
            Text(f.name, fontSize = 12.sp, maxLines = 2, overflow = TextOverflow.Ellipsis, textAlign = androidx.compose.ui.text.style.TextAlign.Center)
        }
        ctxMenu(f, menu) { menu = false }
    }
}

@Composable
private fun BinView(list: List<BinItem>, onRestore: (BinItem) -> Unit, onPurge: (BinItem) -> Unit) {
    if (list.isEmpty()) { Text("Recycle Bin khali hai.", color = W7.gray, modifier = Modifier.padding(24.dp)); return }
    LazyColumn(Modifier.fillMaxSize()) {
        items(list, key = { it.file.absolutePath }) { b ->
            var menu by remember { mutableStateOf(false) }
            Box {
                Row(Modifier.fillMaxWidth().clickable { menu = true }.padding(6.dp), verticalAlignment = Alignment.CenterVertically) {
                    val (ic, tn) = fileIcon(b.file); Icon(ic, null, tint = tn, modifier = Modifier.size(34.dp))
                    Column(Modifier.weight(1f).padding(start = 8.dp)) {
                        Text(b.file.name, fontSize = 14.sp, maxLines = 1, overflow = TextOverflow.Ellipsis)
                        Text("Original: ${File(b.original).parent}", fontSize = 11.sp, color = W7.gray, maxLines = 1, overflow = TextOverflow.Ellipsis)
                        Text("Deleted: ${dateFmt.format(Date(b.time))}", fontSize = 11.sp, color = W7.gray)
                    }
                }
                DropdownMenu(menu, { menu = false }) {
                    DropdownMenuItem({ Text("Restore") }, { menu = false; onRestore(b) }, leadingIcon = { Icon(Icons.Filled.Restore, null) })
                    DropdownMenuItem({ Text("Delete permanently") }, { menu = false; onPurge(b) }, leadingIcon = { Icon(Icons.Filled.DeleteForever, null) })
                }
            }
        }
    }
}

@Composable
fun OpDialog() {
    Dialog({}, DialogProperties(dismissOnBackPress = false, dismissOnClickOutside = false)) {
        Column(Modifier.fillMaxWidth().background(W7.aero).border(1.dp, Color(0xFF5A7FA8)).padding(4.dp)) {
            Text(Op.title, Modifier.padding(6.dp), fontSize = 13.sp)
            Column(Modifier.fillMaxWidth().background(Color.White).padding(14.dp)) {
                Text(Op.title, color = Color(0xFF1E3A5F), fontSize = 16.sp)
                Spacer(Modifier.height(6.dp))
                Text(Op.name, fontSize = 12.sp, maxLines = 1, overflow = TextOverflow.Ellipsis, color = W7.gray)
                Spacer(Modifier.height(10.dp))
                WinProgress(if (Op.total > 0) Op.done.toFloat() / Op.total else null)
                if (Op.total > 0) Text("${fmtSize(Op.done)} of ${fmtSize(Op.total)}", fontSize = 12.sp, modifier = Modifier.padding(top = 6.dp))
            }
        }
    }
}

@Composable
fun PasswordDialog(name: String, wrong: Boolean, onCancel: () -> Unit, onOk: (String) -> Unit) {
    var pw by remember { mutableStateOf("") }
    var show by remember { mutableStateOf(false) }
    WinDialog("Enter password", onCancel, Icons.Filled.Lock, Color(0xFFB8860B), buttons = {
        WinButton("OK", default = true, enabled = pw.isNotEmpty()) { onOk(pw) }
        WinButton("Cancel", onClick = onCancel)
    }) {
        if (wrong) Text("Password galat tha, dobara try karen.", color = Color(0xFFC13A1F), fontSize = 13.sp)
        Text("'$name' ka password likhen:", fontSize = 14.sp)
        Spacer(Modifier.height(6.dp))
        WinField(pw, { pw = it }, password = !show)
        Row(Modifier.clickable { show = !show }.padding(top = 6.dp), verticalAlignment = Alignment.CenterVertically) {
            Checkbox(show, { show = it }); Text("Show password", fontSize = 13.sp)
        }
    }
}

@Composable
private fun PropertiesDialog(files: List<File>, onClose: () -> Unit) {
    var size by remember { mutableLongStateOf(-1L) }
    var cnt by remember { mutableStateOf<Pair<Int, Int>?>(null) }
    LaunchedEffect(files) {
        withContext(Dispatchers.IO) {
            size = Fs.countBytes(files)
            if (files.any { it.isDirectory }) cnt = files.fold(0 to 0) { acc, f -> val c = Fs.counts(f); (acc.first + c.first + (if (f.isFile) 1 else 0)) to (acc.second + c.second) }
        }
    }
    val f = files.first()
    WinDialog(if (files.size == 1) "${f.name} Properties" else "${files.size} items Properties", onClose, fileIcon(f).first, fileIcon(f).second, buttons = {
        WinButton("OK", default = true, onClick = onClose)
    }) {
        @Composable fun line(a: String, b: String) = Row(Modifier.padding(vertical = 3.dp)) {
            Text(a, Modifier.width(92.dp), fontSize = 13.sp, color = W7.gray); Text(b, fontSize = 13.sp)
        }
        if (files.size == 1) {
            line("Name:", f.name)
            line("Type:", Fs.typeName(f))
            line("Location:", f.parent ?: "")
        }
        line("Size:", if (size < 0) "Calculating..." else "${fmtSize(size)} ($size bytes)")
        cnt?.let { line("Contains:", "${it.first} Files, ${it.second} Folders") }
        if (files.size == 1) {
            line("Modified:", dateFmt.format(Date(f.lastModified())))
            line("Attributes:", listOfNotNull(if (!f.canWrite()) "Read-only" else null, if (f.isHidden) "Hidden" else null).joinToString().ifEmpty { "Normal" })
        }
    }
}

@OptIn(ExperimentalFoundationApi::class)
@Composable
fun ImageViewer(list: List<File>, start: Int, onClose: () -> Unit) {
    val ctx = LocalContext.current
    val pager = rememberPagerState(start) { list.size }
    Dialog(onClose, DialogProperties(usePlatformDefaultWidth = false)) {
        Box(Modifier.fillMaxSize().background(Color.Black)) {
            HorizontalPager(pager, Modifier.fillMaxSize()) { i ->
                var zoom by remember { mutableFloatStateOf(1f) }
                AsyncImage(list[i], null, Modifier.fillMaxSize().graphicsLayer(scaleX = zoom, scaleY = zoom)
                    .pointerInput(Unit) { detectTapGestures(onDoubleTap = { zoom = if (zoom > 1f) 1f else 2.5f }) }, contentScale = ContentScale.Fit)
            }
            Row(Modifier.fillMaxWidth().background(Color(0x99000000)).padding(4.dp), verticalAlignment = Alignment.CenterVertically) {
                Text("${pager.currentPage + 1}/${list.size}  ${list.getOrNull(pager.currentPage)?.name ?: ""}", color = Color.White, fontSize = 13.sp, modifier = Modifier.weight(1f).padding(6.dp), maxLines = 1, overflow = TextOverflow.Ellipsis)
                IconButton({ list.getOrNull(pager.currentPage)?.let { Fs.share(ctx, listOf(it)) } }) { Icon(Icons.Filled.Share, null, tint = Color.White) }
                IconButton(onClose) { Icon(Icons.Filled.Close, null, tint = Color.White) }
            }
            Text("Windows Photo Viewer style • double-tap = zoom", color = Color(0x88FFFFFF), fontSize = 11.sp, modifier = Modifier.align(Alignment.BottomCenter).padding(8.dp))
        }
    }
}

/** WinRAR-style archive window: browse inside ZIP/RAR/7Z/TAR, open single files, extract. */
@Composable
fun ArchiveViewer(file: File, opener: Opener, onClose: () -> Unit) {
    val ctx = LocalContext.current
    val scope = rememberCoroutineScope()
    var entries by remember { mutableStateOf<List<AEntry>?>(null) }
    var error by remember { mutableStateOf<String?>(null) }
    var pw by remember { mutableStateOf<String?>(null) }
    var ask by remember { mutableStateOf<Boolean?>(null) }
    var prefix by remember { mutableStateOf("") }
    var busy by remember { mutableStateOf<String?>(null) }

    fun load() {
        scope.launch {
            busy = "Reading..."
            try { entries = withContext(Dispatchers.IO) { Archive.list(file, pw) }; error = null }
            catch (e: NeedPassword) { ask = e.wrong }
            catch (e: Exception) { error = e.message }
            busy = null
        }
    }
    LaunchedEffect(file) { load() }

    fun action(block: suspend () -> String) {
        scope.launch {
            busy = "Extracting..."
            val r = try { withContext(Dispatchers.IO) { block() } } catch (e: NeedPassword) { ask = e.wrong; null } catch (e: Exception) { "Error: ${e.message}" }
            busy = null
            r?.let { Toast.makeText(ctx, it, Toast.LENGTH_LONG).show() }
        }
    }

    // children of current prefix
    val all = entries ?: emptyList()
    val children = remember(entries, prefix) {
        val dirs = LinkedHashSet<String>(); val files = ArrayList<AEntry>()
        all.forEach { e ->
            val n = e.name.trimStart('/')
            if (!n.startsWith(prefix)) return@forEach
            val rest = n.removePrefix(prefix).trimEnd('/')
            if (rest.isEmpty()) return@forEach
            if (rest.contains('/')) dirs += rest.substringBefore('/')
            else if (e.dir) dirs += rest else files += e
        }
        dirs.sorted().map { AEntry(prefix + it + "/", 0, true, 0) } + files.sortedBy { it.name.lowercase() }
    }

    Dialog(onClose, DialogProperties(usePlatformDefaultWidth = false)) {
        WinWindow("${file.name} - FU Archiver", Icons.Filled.Archive, Color(0xFFB8860B), Modifier.padding(6.dp), onClose = onClose) {
            Row(Modifier.fillMaxWidth().background(W7.cmd).horizontalScroll(rememberScrollState()), verticalAlignment = Alignment.CenterVertically) {
                CmdItem("Up", Icons.Filled.ArrowUpward) { prefix = prefix.trimEnd('/').substringBeforeLast('/', "").let { if (it.isEmpty()) "" else "$it/" } }
                CmdItem("Extract All", Icons.Filled.Unarchive, Color(0xFF2E7D32)) {
                    action {
                        val out = Archive.extract(file, file.parentFile ?: Fs.root, null, pw)
                        Fs.scan(ctx, listOf(out.absolutePath))
                        Notify.show(ctx, 9500, "✔ Extract complete", out.name, null)
                        "Extract ho gaya: ${out.absolutePath}"
                    }
                }
                CmdItem("Extract to Downloads", Icons.Filled.Download) {
                    action {
                        val out = Archive.extract(file, Fs.downloads, null, pw)
                        Notify.show(ctx, 9501, "✔ Extract complete", out.name, null)
                        "Extract ho gaya: ${out.absolutePath}"
                    }
                }
            }
            Text("  /$prefix", fontSize = 12.sp, color = W7.gray, modifier = Modifier.fillMaxWidth().background(Color(0xFFF3F6FA)).padding(4.dp))
            Box(Modifier.weight(1f).fillMaxWidth()) {
                error?.let { Text(it, color = Color(0xFFC13A1F), modifier = Modifier.padding(16.dp)) }
                LazyColumn(Modifier.fillMaxSize()) {
                    items(children, key = { it.name }) { e ->
                        val name = e.name.trimEnd('/').substringAfterLast('/')
                        var menu by remember { mutableStateOf(false) }
                        Box {
                            Row(Modifier.fillMaxWidth().clickable {
                                if (e.dir) prefix = e.name.trimStart('/') else action {
                                    val f = Archive.extractToCache(file, e.name, ctx.cacheDir, pw)
                                    withContext(Dispatchers.Main) { opener.open(f) }; "Opened ${f.name}"
                                }
                            }.padding(8.dp), verticalAlignment = Alignment.CenterVertically) {
                                val (ic, tn) = if (e.dir) Icons.Filled.Folder to W7.folder else fileIcon(File(name))
                                Icon(ic, null, tint = tn, modifier = Modifier.size(30.dp))
                                Column(Modifier.weight(1f).padding(start = 8.dp)) {
                                    Text(name, fontSize = 14.sp, maxLines = 1, overflow = TextOverflow.Ellipsis)
                                    if (!e.dir) Text((if (e.size >= 0) fmtSize(e.size) else "") + (if (e.time > 0) "   " + dateFmt.format(Date(e.time)) else ""), fontSize = 11.sp, color = W7.gray)
                                }
                                if (!e.dir) IconButton({ menu = true }) { Icon(Icons.Filled.MoreVert, null) }
                            }
                            DropdownMenu(menu, { menu = false }) {
                                DropdownMenuItem({ Text("Open") }, { menu = false; action { val f = Archive.extractToCache(file, e.name, ctx.cacheDir, pw); withContext(Dispatchers.Main) { opener.open(f) }; "Opened" } })
                                DropdownMenuItem({ Text("Extract this file here") }, { menu = false; action {
                                    Archive.extract(file, file.parentFile ?: Fs.root, setOf(e.name), pw); "$name extract ho gayi"
                                } })
                            }
                        }
                    }
                }
                busy?.let { b ->
                    Column(Modifier.align(Alignment.Center).background(Color.White).border(1.dp, W7.border).padding(16.dp)) {
                        Text(b); Spacer(Modifier.height(8.dp)); WinProgress(null, Modifier.width(200.dp))
                    }
                }
            }
            Text("  ${all.count { !it.dir }} files  •  ${fmtSize(file.length())} archive${if (pw != null) "  •  🔒 unlocked" else ""}", fontSize = 12.sp,
                modifier = Modifier.fillMaxWidth().background(Color(0xFFDCE6F3)).padding(5.dp))
        }
    }

    ask?.let { wrong ->
        PasswordDialog(file.name, wrong, { ask = null; if (entries == null) onClose() }) { p -> ask = null; pw = p; if (entries == null) load() }
    }
}
