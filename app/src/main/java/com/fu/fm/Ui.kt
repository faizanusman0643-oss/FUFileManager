@file:OptIn(ExperimentalFoundationApi::class, ExperimentalMaterial3Api::class)

package com.fu.fm

import android.content.Context
import android.content.Intent
import android.net.Uri
import android.webkit.MimeTypeMap
import android.widget.Toast
import androidx.activity.compose.BackHandler
import androidx.compose.foundation.ExperimentalFoundationApi
import androidx.compose.foundation.background
import androidx.compose.foundation.combinedClickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.LazyRow
import androidx.compose.foundation.lazy.grid.GridCells
import androidx.compose.foundation.lazy.grid.LazyVerticalGrid
import androidx.compose.foundation.lazy.grid.itemsIndexed
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.vector.ImageVector
import androidx.compose.ui.layout.ContentScale
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.text.input.PasswordVisualTransformation
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.core.content.FileProvider
import coil.ImageLoader
import coil.compose.AsyncImage
import java.io.File

val Teal = Color(0xFF0E7C86)
val Navy = Color(0xFF0B2545)
private val TealLight = Color(0xFFD5F0F2)

@Composable
fun FuTheme(content: @Composable () -> Unit) {
    MaterialTheme(
        colorScheme = lightColorScheme(primary = Teal, secondary = Navy, primaryContainer = TealLight),
        content = content
    )
}

fun iconFor(c: Cat): ImageVector = when (c) {
    Cat.AUDIO -> Icons.Default.MusicNote
    Cat.DOC -> Icons.Default.Description
    Cat.PHOTO -> Icons.Default.Image
    Cat.VIDEO -> Icons.Default.Movie
    Cat.OTHER -> Icons.Default.InsertDriveFile
}

fun openFile(ctx: Context, f: File) {
    try {
        val uri = FileProvider.getUriForFile(ctx, ctx.packageName + ".fp", f)
        val mime = MimeTypeMap.getSingleton().getMimeTypeFromExtension(f.extension.lowercase()) ?: "*/*"
        val i = Intent(Intent.ACTION_VIEW).setDataAndType(uri, mime)
            .addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_ACTIVITY_NEW_TASK)
        ctx.startActivity(i)
    } catch (e: Exception) {
        Toast.makeText(ctx, "Is file ko kholne wali app nahi mili", Toast.LENGTH_SHORT).show()
    }
}

fun shareFiles(ctx: Context, files: List<File>) {
    val uris = ArrayList<Uri>(files.filter { it.isFile }.map { FileProvider.getUriForFile(ctx, ctx.packageName + ".fp", it) })
    if (uris.isEmpty()) return
    val i = Intent(Intent.ACTION_SEND_MULTIPLE).setType("*/*")
        .putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris)
        .addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    ctx.startActivity(Intent.createChooser(i, "Share").addFlags(Intent.FLAG_ACTIVITY_NEW_TASK))
}

@Composable
fun App(vm: VM, loader: ImageLoader, extUri: Uri?, onExtClose: () -> Unit) {
    val ctx = LocalContext.current
    var tab by remember { mutableIntStateOf(0) }
    var player by remember { mutableStateOf<Pair<List<Item>, Int>?>(null) }
    var viewer by remember { mutableStateOf<Pair<List<Item>, Int>?>(null) }
    var picker by remember { mutableStateOf<Pair<String, (File) -> Unit>?>(null) }
    var confirm by remember { mutableStateOf<Pair<String, () -> Unit>?>(null) }
    var visible by remember { mutableStateOf<List<String>>(emptyList()) }
    var menu by remember { mutableStateOf(false) }
    val snack = remember { SnackbarHostState() }
    val sel = vm.selected

    LaunchedEffect(vm.message) {
        vm.message?.let { snack.showSnackbar(it); vm.message = null }
    }
    BackHandler(sel.isNotEmpty()) { sel.clear() }

    fun openItem(list: List<Item>, item: Item) {
        when (item.cat) {
            Cat.PHOTO -> {
                val l = list.filter { it.cat == Cat.PHOTO }
                viewer = l to l.indexOfFirst { it.path == item.path }
            }
            Cat.VIDEO, Cat.AUDIO -> {
                val l = list.filter { it.cat == item.cat }
                player = l to l.indexOfFirst { it.path == item.path }
            }
            else -> openFile(ctx, item.file)
        }
    }

    Box(Modifier.fillMaxSize()) {
        Scaffold(
            topBar = {
                if (sel.isNotEmpty()) {
                    TopAppBar(
                        title = { Text("${sel.size} selected") },
                        navigationIcon = { IconButton(onClick = { sel.clear() }) { Icon(Icons.Default.Close, "Clear") } },
                        actions = {
                            IconButton(onClick = { sel.clear(); sel.addAll(visible) }) { Icon(Icons.Default.SelectAll, "Select all") }
                            IconButton(onClick = {
                                val l = sel.toList()
                                picker = "Move to…" to { d: File -> vm.move(l, d, false) }
                            }) { Icon(Icons.Default.DriveFileMove, "Move") }
                            IconButton(onClick = {
                                val l = sel.toList()
                                confirm = "${l.size} item hamesha ke liye delete karni hain?" to { vm.delete(l) }
                            }) { Icon(Icons.Default.Delete, "Delete") }
                            Box {
                                IconButton(onClick = { menu = true }) { Icon(Icons.Default.MoreVert, "More") }
                                DropdownMenu(expanded = menu, onDismissRequest = { menu = false }) {
                                    DropdownMenuItem(text = { Text("Copy to…") }, onClick = {
                                        menu = false
                                        val l = sel.toList()
                                        picker = "Copy to…" to { d: File -> vm.move(l, d, true) }
                                    })
                                    DropdownMenuItem(text = { Text("Share") }, onClick = {
                                        menu = false
                                        shareFiles(ctx, sel.map { File(it) })
                                    })
                                    DropdownMenuItem(text = { Text("Vault mai chhupao") }, onClick = {
                                        menu = false
                                        vm.hide(sel.toList())
                                    })
                                }
                            }
                        }
                    )
                } else {
                    TopAppBar(
                        title = { Text("FU File Manager") },
                        colors = TopAppBarDefaults.topAppBarColors(containerColor = Navy, titleContentColor = Color.White)
                    )
                }
            },
            bottomBar = {
                NavigationBar {
                    listOf(
                        "Gallery" to Icons.Default.PhotoLibrary,
                        "Files" to Icons.Default.Folder,
                        "Tools" to Icons.Default.Build,
                        "Vault" to Icons.Default.Lock
                    ).forEachIndexed { i, (t, ic) ->
                        NavigationBarItem(
                            selected = tab == i,
                            onClick = { tab = i; sel.clear(); vm.vaultOpen = false },
                            icon = { Icon(ic, t) },
                            label = { Text(t) }
                        )
                    }
                }
            },
            snackbarHost = { SnackbarHost(snack) }
        ) { pad ->
            Box(Modifier.padding(pad).fillMaxSize()) {
                when (tab) {
                    0 -> Gallery(vm, loader, { l, x -> openItem(l, x) }) { v -> visible = v }
                    1 -> Browser(vm, { l, x -> openItem(l, x) }) { v -> visible = v }
                    2 -> Tools(vm, { t, cb -> picker = t to cb }, { m, a -> confirm = m to a })
                    else -> VaultScreen(vm)
                }
                if (vm.busy) {
                    Box(
                        Modifier.fillMaxSize().background(Color(0x66000000))
                            .combinedClickable(onClick = {}),
                        Alignment.Center
                    ) {
                        Card(Modifier.padding(32.dp).fillMaxWidth()) {
                            Column(Modifier.padding(20.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
                                Text("Kaam ho raha hai…", fontSize = 16.sp)
                                if (vm.progTotal > 0) {
                                    LinearProgressIndicator(
                                        progress = { vm.progDone.toFloat() / vm.progTotal },
                                        modifier = Modifier.fillMaxWidth()
                                    )
                                    Text("${vm.progDone} / ${vm.progTotal}", fontSize = 13.sp)
                                } else {
                                    LinearProgressIndicator(Modifier.fillMaxWidth())
                                }
                                if (vm.progName.isNotEmpty()) {
                                    Text(vm.progName, fontSize = 12.sp, color = Color.Gray, maxLines = 1, overflow = TextOverflow.Ellipsis)
                                }
                            }
                        }
                    }
                }
            }
        }

        player?.let { (l, i) ->
            PlayerScreen(l.map { Uri.fromFile(it.file) }, l.map { it.name }, i, l[i].cat == Cat.AUDIO) { player = null }
        }
        viewer?.let { (l, i) -> ImageViewer(l, i, loader) { viewer = null } }
        if (extUri != null) {
            val isAudio = ctx.contentResolver.getType(extUri)?.startsWith("audio") == true
            PlayerScreen(listOf(extUri), listOf(extUri.lastPathSegment ?: "Media"), 0, isAudio, onExtClose)
        }
        vm.dupes?.let { g ->
            DupesDialog(g, onDelete = { p -> vm.dupes = null; vm.delete(p) }, onClose = { vm.dupes = null })
        }
    }

    picker?.let { (title, _) ->
        key(title) {
            FolderPicker(
                title,
                onPick = { f ->
                    val cb = picker?.second
                    picker = null
                    cb?.invoke(f)
                },
                onDismiss = { picker = null }
            )
        }
    }
    confirm?.let { (msg, act) ->
        AlertDialog(
            onDismissRequest = { confirm = null },
            title = { Text("Confirm") },
            text = { Text(msg) },
            confirmButton = { TextButton(onClick = { confirm = null; act() }) { Text("Haan") } },
            dismissButton = { TextButton(onClick = { confirm = null }) { Text("Nahi") } }
        )
    }
}

@Composable
fun Gallery(
    vm: VM,
    loader: ImageLoader,
    open: (List<Item>, Item) -> Unit,
    onVisible: (List<String>) -> Unit
) {
    val list = remember(vm.items, vm.filter) { vm.items.filter { vm.filter == null || it.cat == vm.filter } }
    val counts = remember(vm.items) { vm.items.groupingBy { it.cat }.eachCount() }
    LaunchedEffect(list) { onVisible(list.map { it.path }) }
    Column(Modifier.fillMaxSize()) {
        LazyRow(contentPadding = PaddingValues(8.dp), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
            item {
                FilterChip(selected = vm.filter == null, onClick = { vm.filter = null },
                    label = { Text("All (${vm.items.size})") })
            }
            items(Cat.values().toList()) { c ->
                FilterChip(selected = vm.filter == c, onClick = { vm.filter = c },
                    label = { Text("${c.label} (${counts[c] ?: 0})") })
            }
        }
        if (vm.loading && vm.items.isEmpty()) {
            Box(Modifier.fillMaxSize(), Alignment.Center) { CircularProgressIndicator() }
        } else if (list.isEmpty()) {
            Box(Modifier.fillMaxSize(), Alignment.Center) { Text("Kuch nahi mila", color = Color.Gray) }
        } else {
            LazyVerticalGrid(
                columns = GridCells.Adaptive(104.dp),
                contentPadding = PaddingValues(2.dp),
                horizontalArrangement = Arrangement.spacedBy(2.dp),
                verticalArrangement = Arrangement.spacedBy(2.dp)
            ) {
                itemsIndexed(list, key = { _, e -> e.path }) { _, e ->
                    Cell(
                        e, e.path in vm.selected, loader,
                        onClick = { if (vm.selected.isNotEmpty()) vm.toggle(e.path) else open(list, e) },
                        onLong = { vm.toggle(e.path) }
                    )
                }
            }
        }
    }
}

@Composable
fun Cell(item: Item, selected: Boolean, loader: ImageLoader, onClick: () -> Unit, onLong: () -> Unit) {
    Box(
        Modifier.aspectRatio(1f).background(Color(0xFFE8EEF2))
            .combinedClickable(onClick = onClick, onLongClick = onLong)
    ) {
        when (item.cat) {
            Cat.PHOTO, Cat.VIDEO -> AsyncImage(
                model = item.file, imageLoader = loader, contentDescription = null,
                contentScale = ContentScale.Crop, modifier = Modifier.fillMaxSize()
            )
            else -> Column(
                Modifier.fillMaxSize().padding(6.dp),
                Arrangement.Center, Alignment.CenterHorizontally
            ) {
                Icon(iconFor(item.cat), null, Modifier.size(36.dp), tint = Teal)
                Text(item.name, maxLines = 2, overflow = TextOverflow.Ellipsis, fontSize = 11.sp, textAlign = TextAlign.Center)
            }
        }
        if (item.cat == Cat.VIDEO) {
            Icon(Icons.Default.PlayCircle, null, Modifier.align(Alignment.BottomStart).padding(4.dp).size(22.dp), tint = Color.White)
        }
        if (selected) {
            Box(Modifier.fillMaxSize().background(Teal.copy(alpha = 0.35f)))
            Icon(Icons.Default.CheckCircle, null, Modifier.align(Alignment.TopEnd).padding(4.dp), tint = Color.White)
        }
    }
}

@Composable
fun Browser(vm: VM, open: (List<Item>, Item) -> Unit, onVisible: (List<String>) -> Unit) {
    var dir by remember { mutableStateOf(Fm.root) }
    val entries = remember(dir, vm.version) {
        (dir.listFiles()?.filter { !it.name.startsWith(".") } ?: emptyList())
            .sortedWith(compareBy({ !it.isDirectory }, { it.name.lowercase() }))
    }
    LaunchedEffect(entries) { onVisible(entries.map { it.absolutePath }) }
    val atRoot = dir.absolutePath == Fm.root.absolutePath
    BackHandler(!atRoot && vm.selected.isEmpty()) { dir = dir.parentFile ?: Fm.root }
    Column(Modifier.fillMaxSize()) {
        Text(
            dir.absolutePath.removePrefix(Fm.root.absolutePath).ifEmpty { "/" },
            Modifier.padding(12.dp, 8.dp), fontSize = 12.sp, color = Color.Gray
        )
        LazyColumn(Modifier.fillMaxSize()) {
            items(entries, key = { it.absolutePath }) { f ->
                val isSel = f.absolutePath in vm.selected
                ListItem(
                    modifier = Modifier.combinedClickable(
                        onClick = {
                            if (vm.selected.isNotEmpty()) vm.toggle(f.absolutePath)
                            else if (f.isDirectory) dir = f
                            else {
                                val files = entries.filter { e -> e.isFile }.map { e -> e.toItem() }
                                open(files, f.toItem())
                            }
                        },
                        onLongClick = { vm.toggle(f.absolutePath) }
                    ),
                    colors = ListItemDefaults.colors(containerColor = if (isSel) TealLight else Color.Transparent),
                    leadingContent = {
                        Icon(
                            if (f.isDirectory) Icons.Default.Folder else iconFor(catOf(f.name)), null,
                            tint = if (f.isDirectory) Color(0xFFE0A100) else Teal
                        )
                    },
                    headlineContent = { Text(f.name, maxLines = 1, overflow = TextOverflow.Ellipsis) },
                    supportingContent = {
                        Text(if (f.isDirectory) "Folder" else "${fmtSize(f.length())} • ${fmtDate(f.lastModified())}")
                    }
                )
            }
        }
    }
}

@Composable
fun Tools(vm: VM, pick: (String, (File) -> Unit) -> Unit, confirm: (String, () -> Unit) -> Unit) {
    var modeDir by remember { mutableStateOf<File?>(null) }
    Column(
        Modifier.fillMaxSize().verticalScroll(rememberScrollState()).padding(12.dp),
        verticalArrangement = Arrangement.spacedBy(10.dp)
    ) {
        ToolCard("Duplicate finder", "Same file alag naam se — dhoond kar delete karo", Icons.Default.ContentCopy) {
            vm.findDupes()
        }
        ToolCard("Sort by Format / Date / Month", "Folder ki files subfolders mai arrange", Icons.Default.SortByAlpha) {
            pick("Sort karne wala folder") { modeDir = it }
        }
        ToolCard("Consolidate (Reverse Data)", "Sab subfolders ki files ek folder mai", Icons.Default.CallMerge) {
            pick("Source folder (is ke andar ki sab files)") { src ->
                pick("Destination folder") { dst ->
                    confirm("'${src.name}' ki sab files '${dst.name}' mai move hongi. Theek hai?") { vm.consolidate(src, dst) }
                }
            }
        }
        ToolCard("Delete empty folders", "Khali folders hata do", Icons.Default.DeleteSweep) {
            pick("Folder chuno") { d ->
                confirm("'${d.name}' ke andar ke khali folders delete honge.") { vm.emptyDirs(d) }
            }
        }
        ToolCard("Clean junk", ".tmp files, thumbnails, khali folders", Icons.Default.CleaningServices) {
            confirm("Temp aur thumbnail files delete hongi. Theek hai?") { vm.clean() }
        }
    }
    modeDir?.let { d ->
        AlertDialog(
            onDismissRequest = { modeDir = null },
            title = { Text("Sort: ${d.name}") },
            text = {
                Column {
                    listOf(
                        "By format (extension)" to SortMode.FORMAT,
                        "By date" to SortMode.DATE,
                        "By month" to SortMode.MONTH
                    ).forEach { (t, m) ->
                        TextButton(onClick = { modeDir = null; vm.sort(d, m) }) { Text(t) }
                    }
                }
            },
            confirmButton = {},
            dismissButton = { TextButton(onClick = { modeDir = null }) { Text("Cancel") } }
        )
    }
}

@Composable
fun ToolCard(title: String, sub: String, icon: ImageVector, onClick: () -> Unit) {
    Card(onClick = onClick, Modifier.fillMaxWidth()) {
        Row(Modifier.padding(16.dp), verticalAlignment = Alignment.CenterVertically) {
            Icon(icon, null, tint = Teal, modifier = Modifier.size(28.dp))
            Spacer(Modifier.width(16.dp))
            Column {
                Text(title, fontSize = 16.sp)
                Text(sub, fontSize = 12.sp, color = Color.Gray)
            }
        }
    }
}

@Composable
fun VaultScreen(vm: VM) {
    val ctx = LocalContext.current
    var pin by remember { mutableStateOf("") }
    var pin2 by remember { mutableStateOf("") }
    var err by remember { mutableStateOf<String?>(null) }
    var isSet by remember { mutableStateOf(Pin.isSet(ctx)) }
    val chosen = remember { mutableStateListOf<String>() }
    val kb = KeyboardOptions(keyboardType = KeyboardType.NumberPassword)

    Column(Modifier.fillMaxSize().padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
        if (!isSet) {
            Text("Vault ka PIN set karo (kam az kam 4 digits)")
            OutlinedTextField(pin, { pin = it }, label = { Text("PIN") }, singleLine = true,
                visualTransformation = PasswordVisualTransformation(), keyboardOptions = kb)
            OutlinedTextField(pin2, { pin2 = it }, label = { Text("PIN dobara") }, singleLine = true,
                visualTransformation = PasswordVisualTransformation(), keyboardOptions = kb)
            Button(onClick = {
                if (pin.length < 4) err = "PIN kam az kam 4 digits ka ho"
                else if (pin != pin2) err = "PIN match nahi hua"
                else {
                    Pin.set(ctx, pin); isSet = true; vm.vaultOpen = true
                    pin = ""; pin2 = ""; err = null
                }
            }) { Text("Set PIN") }
        } else if (!vm.vaultOpen) {
            Text("Vault locked hai")
            OutlinedTextField(pin, { pin = it }, label = { Text("PIN") }, singleLine = true,
                visualTransformation = PasswordVisualTransformation(), keyboardOptions = kb)
            Button(onClick = {
                if (Pin.check(ctx, pin)) { vm.vaultOpen = true; pin = ""; err = null } else err = "Galat PIN"
            }) { Text("Unlock") }
        } else {
            val entries = remember(vm.version) { Vault.entries() }
            Row(verticalAlignment = Alignment.CenterVertically) {
                Text("Vault (${entries.size})", Modifier.weight(1f), fontSize = 18.sp)
                TextButton(onClick = { vm.vaultOpen = false }) { Text("Lock") }
            }
            Text("Chhupane ke liye Gallery/Files mai select karo, phir ⋮ menu > Vault mai chhupao", fontSize = 12.sp, color = Color.Gray)
            LazyColumn(Modifier.weight(1f)) {
                items(entries, key = { it.first }) { e ->
                    ListItem(
                        leadingContent = {
                            Checkbox(checked = e.first in chosen, onCheckedChange = { c ->
                                if (c) chosen.add(e.first) else chosen.remove(e.first)
                            })
                        },
                        headlineContent = { Text(e.first, maxLines = 1, overflow = TextOverflow.Ellipsis) },
                        supportingContent = { Text(e.second.removePrefix(Fm.root.absolutePath), fontSize = 11.sp) }
                    )
                }
            }
            Button(enabled = chosen.isNotEmpty(), onClick = { vm.restore(chosen.toList()); chosen.clear() }) {
                Text("Unhide ${chosen.size}")
            }
        }
        err?.let { Text(it, color = MaterialTheme.colorScheme.error) }
    }
}
