@file:OptIn(ExperimentalFoundationApi::class, ExperimentalMaterial3Api::class)

package com.fu.fm

import android.Manifest
import android.os.Build
import androidx.activity.compose.BackHandler
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.ExperimentalFoundationApi
import androidx.compose.foundation.background
import androidx.compose.foundation.combinedClickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyRow
import androidx.compose.foundation.lazy.grid.GridCells
import androidx.compose.foundation.lazy.grid.LazyVerticalGrid
import androidx.compose.foundation.lazy.grid.items
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import coil.ImageLoader

private enum class RecTab(val label: String) { DELETED("Deleted"), STATUS("Status"), ALL("Sab backup") }

@Composable
fun BusyOverlay(vm: VM) {
    if (!vm.busy) return
    Box(
        Modifier.fillMaxSize().background(Color(0x66000000)).combinedClickable(onClick = {}),
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

@Composable
fun RecoveryScreen(
    vm: VM,
    loader: ImageLoader,
    open: (List<Item>, Item) -> Unit,
    confirm: (String, () -> Unit) -> Unit,
    onClose: () -> Unit
) {
    val sel = vm.selected
    var tab by remember { mutableStateOf(RecTab.DELETED) }
    var menu by remember { mutableStateOf(false) }
    val snack = remember { SnackbarHostState() }
    LaunchedEffect(vm.message) { vm.message?.let { snack.showSnackbar(it); vm.message = null } }
    val notifPerm = rememberLauncherForActivityResult(ActivityResultContracts.RequestPermission()) { }

    LaunchedEffect(Unit) { sel.clear(); vm.loadRecs(vm.waAuto) }
    BackHandler { if (sel.isNotEmpty()) sel.clear() else onClose() }

    val shown = remember(vm.recs, tab) {
        when (tab) {
            RecTab.DELETED -> vm.recs.filter { it.deleted }
            RecTab.STATUS -> vm.recs.filter { it.status }
            RecTab.ALL -> vm.recs.filter { !it.trashed }
        }
    }
    val counts = remember(vm.recs) {
        mapOf(
            RecTab.DELETED to vm.recs.count { it.deleted },
            RecTab.STATUS to vm.recs.count { it.status },
            RecTab.ALL to vm.recs.count { !it.trashed }
        )
    }

    Box(Modifier.fillMaxSize().background(MaterialTheme.colorScheme.background)) {
        Scaffold(
            snackbarHost = { SnackbarHost(snack) },
            topBar = {
                if (sel.isNotEmpty()) {
                    TopAppBar(
                        title = { Text("${sel.size} selected") },
                        navigationIcon = { IconButton(onClick = { sel.clear() }) { Icon(Icons.Default.Close, "Clear") } },
                        actions = {
                            IconButton(onClick = { sel.clear(); sel.addAll(shown.map { it.item.path }) }) {
                                Icon(Icons.Default.SelectAll, "Select all")
                            }
                            IconButton(onClick = { vm.recover(sel.toList()) }) { Icon(Icons.Default.Restore, "Recover") }
                            IconButton(onClick = {
                                val l = sel.toList()
                                confirm("${l.size} backup/trash files hamesha ke liye delete karni hain? Phir wapas nahi aayengi.") {
                                    vm.deleteRecs(l)
                                }
                            }) { Icon(Icons.Default.Delete, "Delete") }
                        }
                    )
                } else {
                    TopAppBar(
                        title = { Text("WhatsApp Recovery") },
                        navigationIcon = {
                            IconButton(onClick = onClose) { Icon(Icons.AutoMirrored.Filled.ArrowBack, "Back", tint = Color.White) }
                        },
                        actions = {
                            IconButton(onClick = { vm.loadRecs(true) }) { Icon(Icons.Default.Sync, "Backup now", tint = Color.White) }
                            Box {
                                IconButton(onClick = { menu = true }) { Icon(Icons.Default.MoreVert, "More", tint = Color.White) }
                                DropdownMenu(expanded = menu, onDismissRequest = { menu = false }) {
                                    DropdownMenuItem(text = { Text("Purana backup saaf karo (30 din+)") }, onClick = {
                                        menu = false
                                        confirm("30 din se purani copies hatengi jin ki original file abhi phone mai maujood hai. Deleted files ki copies nahi hatengi.") {
                                            vm.pruneRecs(30)
                                        }
                                    })
                                }
                            }
                        },
                        colors = TopAppBarDefaults.topAppBarColors(containerColor = Navy, titleContentColor = Color.White)
                    )
                }
            }
        ) { pad ->
            Column(Modifier.padding(pad).fillMaxSize()) {
                Card(Modifier.fillMaxWidth().padding(10.dp)) {
                    Row(Modifier.padding(14.dp), verticalAlignment = Alignment.CenterVertically) {
                        Column(Modifier.weight(1f)) {
                            Text("Auto backup", fontSize = 16.sp)
                            Text(
                                if (vm.waAuto) "ON — har nayi WhatsApp photo/video/voice/status ki copy ban rahi hai. " +
                                    "Backup: ${counts[RecTab.ALL]} files • ${fmtSize(WaRecovery.backupSize(vm.recs))}"
                                else "OFF — ON karo taake aage se delete hone wali files bach jayen",
                                fontSize = 12.sp, color = Color.Gray
                            )
                        }
                        Switch(checked = vm.waAuto, onCheckedChange = { on ->
                            if (on && Build.VERSION.SDK_INT >= 33) notifPerm.launch(Manifest.permission.POST_NOTIFICATIONS)
                            vm.setWaAuto(on)
                        })
                    }
                }
                LazyRow(contentPadding = PaddingValues(horizontal = 10.dp), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                    items(RecTab.values().size) { i ->
                        val t = RecTab.values()[i]
                        FilterChip(selected = tab == t, onClick = { tab = t; sel.clear() },
                            label = { Text("${t.label} (${counts[t] ?: 0})") })
                    }
                }
                if (shown.isEmpty()) {
                    Box(Modifier.fillMaxSize().padding(24.dp), Alignment.Center) {
                        Text(
                            when (tab) {
                                RecTab.DELETED ->
                                    "Abhi koi deleted file nahi mili.\n\nJab WhatsApp pe koi photo/video delete hogi " +
                                        "(\"Delete for everyone\" ya aap khud delete karo), wo yahan aa jayegi — " +
                                        "bas Auto backup ON rakho.\n\nNote: Auto backup ON karne se PEHLE delete hui files " +
                                        "root ke baghair wapas nahi aa saktin (sirf Android trash wali milti hain)."
                                RecTab.STATUS -> "Jo status aap dekhte ho un ki copy yahan save hoti hai (24 ghante baad bhi)."
                                RecTab.ALL -> "Backup khali hai. Auto backup ON karo ya upar ⟳ dabao."
                            },
                            textAlign = TextAlign.Center, color = Color.Gray
                        )
                    }
                } else {
                    Text(
                        "Long-press se select karo, phir ⟲ Recover — files 'Download/WhatsApp Recovered' mai jayengi",
                        Modifier.padding(10.dp, 6.dp), fontSize = 11.sp, color = Color.Gray
                    )
                    LazyVerticalGrid(
                        columns = GridCells.Adaptive(104.dp),
                        contentPadding = PaddingValues(2.dp),
                        horizontalArrangement = Arrangement.spacedBy(2.dp),
                        verticalArrangement = Arrangement.spacedBy(2.dp)
                    ) {
                        items(shown, key = { it.item.path }) { r ->
                            Box {
                                Cell(
                                    r.item, r.item.path in sel, loader,
                                    onClick = {
                                        if (sel.isNotEmpty()) vm.toggle(r.item.path)
                                        else open(shown.map { it.item }, r.item)
                                    },
                                    onLong = { vm.toggle(r.item.path) }
                                )
                                if (r.deleted && tab != RecTab.DELETED) {
                                    Text(
                                        "DELETED", Modifier.align(Alignment.TopStart).background(Color(0xCCB00020))
                                            .padding(horizontal = 4.dp), color = Color.White, fontSize = 9.sp
                                    )
                                }
                            }
                        }
                    }
                }
            }
        }
        BusyOverlay(vm)
    }
}
