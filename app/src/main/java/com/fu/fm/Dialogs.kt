@file:OptIn(ExperimentalMaterial3Api::class)

package com.fu.fm

import android.net.Uri
import androidx.activity.compose.BackHandler
import androidx.compose.foundation.ExperimentalFoundationApi
import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.pager.HorizontalPager
import androidx.compose.foundation.pager.rememberPagerState
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Close
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.layout.ContentScale
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.compose.ui.viewinterop.AndroidView
import androidx.compose.ui.window.Dialog
import androidx.compose.ui.window.DialogProperties
import androidx.media3.common.MediaItem
import androidx.media3.common.Player
import androidx.media3.exoplayer.ExoPlayer
import androidx.media3.ui.PlayerView
import coil.ImageLoader
import coil.compose.AsyncImage
import java.io.File

@Composable
fun FolderPicker(title: String, onPick: (File) -> Unit, onDismiss: () -> Unit) {
    var dir by remember { mutableStateOf(Fm.root) }
    var newName by remember { mutableStateOf("") }
    var tick by remember { mutableIntStateOf(0) }
    val subs = remember(dir, tick) {
        (dir.listFiles()?.filter { it.isDirectory && !it.name.startsWith(".") } ?: emptyList())
            .sortedBy { it.name.lowercase() }
    }
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(title, fontSize = 16.sp) },
        text = {
            Column {
                Text(dir.absolutePath, fontSize = 11.sp, color = Color.Gray)
                LazyColumn(Modifier.height(260.dp)) {
                    if (dir.absolutePath != Fm.root.absolutePath) {
                        item {
                            TextButton(onClick = { dir = dir.parentFile ?: Fm.root }, modifier = Modifier.fillMaxWidth()) {
                                Text("⬆  upar (..)", Modifier.fillMaxWidth(), textAlign = TextAlign.Start)
                            }
                        }
                    }
                    items(subs, key = { it.absolutePath }) { f ->
                        TextButton(onClick = { dir = f }, modifier = Modifier.fillMaxWidth()) {
                            Text(f.name, Modifier.fillMaxWidth(), textAlign = TextAlign.Start, maxLines = 1)
                        }
                    }
                }
                Row(verticalAlignment = Alignment.CenterVertically) {
                    OutlinedTextField(
                        value = newName, onValueChange = { newName = it }, modifier = Modifier.weight(1f),
                        placeholder = { Text("Naya folder") }, singleLine = true
                    )
                    TextButton(onClick = {
                        if (newName.isNotBlank()) { File(dir, newName.trim()).mkdirs(); newName = ""; tick++ }
                    }) { Text("Banao") }
                }
            }
        },
        confirmButton = { TextButton(onClick = { onPick(dir) }) { Text("Yahan select karo") } },
        dismissButton = { TextButton(onClick = onDismiss) { Text("Cancel") } }
    )
}

@Composable
fun DupesDialog(groups: List<List<Item>>, onDelete: (List<String>) -> Unit, onClose: () -> Unit) {
    val checked = remember {
        mutableStateListOf<String>().apply { groups.forEach { g -> g.drop(1).forEach { add(it.path) } } }
    }
    val total = groups.flatten().filter { it.path in checked }.sumOf { it.size }
    Dialog(onDismissRequest = onClose, properties = DialogProperties(usePlatformDefaultWidth = false)) {
        Surface(Modifier.fillMaxSize()) {
            Column(Modifier.fillMaxSize()) {
                TopAppBar(
                    title = { Text("${groups.size} duplicate groups") },
                    navigationIcon = { IconButton(onClick = onClose) { Icon(Icons.Default.Close, "Close") } }
                )
                Text(
                    "Har group mai sabse purani file bachti hai. Baaki tick hain, chaho to badal lo.",
                    Modifier.padding(12.dp, 4.dp), fontSize = 12.sp, color = Color.Gray
                )
                LazyColumn(Modifier.weight(1f)) {
                    groups.forEachIndexed { gi, g ->
                        item(key = "h$gi") {
                            Text(
                                "${g[0].name} • ${fmtSize(g[0].size)}",
                                Modifier.padding(12.dp, 12.dp, 12.dp, 2.dp), fontWeight = FontWeight.Bold
                            )
                        }
                        items(g, key = { it.path }) { e ->
                            Row(Modifier.fillMaxWidth(), verticalAlignment = Alignment.CenterVertically) {
                                Checkbox(checked = e.path in checked, onCheckedChange = { c ->
                                    if (c) {
                                        if (g.count { x -> x.path in checked } < g.size - 1) checked.add(e.path)
                                    } else checked.remove(e.path)
                                })
                                Column {
                                    Text(e.path.removePrefix(Fm.root.absolutePath), fontSize = 12.sp)
                                    Text(fmtDate(e.mtime), fontSize = 11.sp, color = Color.Gray)
                                }
                            }
                        }
                    }
                }
                Button(
                    onClick = { onDelete(checked.toList()) },
                    enabled = checked.isNotEmpty(),
                    modifier = Modifier.fillMaxWidth().padding(12.dp)
                ) { Text("Delete ${checked.size} files (${fmtSize(total)})") }
            }
        }
    }
}

@Composable
fun PlayerScreen(uris: List<Uri>, names: List<String>, start: Int, audio: Boolean, onClose: () -> Unit) {
    val ctx = LocalContext.current
    val exo = remember {
        ExoPlayer.Builder(ctx).build().apply {
            setMediaItems(uris.map { MediaItem.fromUri(it) })
            seekTo(start, 0L)
            prepare()
            playWhenReady = true
        }
    }
    var title by remember { mutableStateOf(names.getOrElse(start) { "" }) }
    DisposableEffect(exo) {
        val l = object : Player.Listener {
            override fun onMediaItemTransition(mediaItem: MediaItem?, reason: Int) {
                title = names.getOrElse(exo.currentMediaItemIndex) { "" }
            }
        }
        exo.addListener(l)
        onDispose { exo.removeListener(l); exo.release() }
    }
    BackHandler(onBack = onClose)
    Box(Modifier.fillMaxSize().background(Color.Black)) {
        AndroidView(
            factory = { c ->
                PlayerView(c).apply {
                    player = exo
                    useController = true
                    controllerShowTimeoutMs = if (audio) 0 else 3000
                    if (audio) showController()
                }
            },
            modifier = Modifier.fillMaxSize()
        )
        IconButton(onClick = onClose, modifier = Modifier.align(Alignment.TopStart).padding(8.dp)) {
            Icon(Icons.AutoMirrored.Filled.ArrowBack, "Back", tint = Color.White)
        }
        Text(
            title, color = Color.White, fontSize = 13.sp, maxLines = 1,
            modifier = Modifier.align(Alignment.TopCenter).padding(top = 18.dp, start = 56.dp, end = 16.dp)
        )
    }
}

@OptIn(ExperimentalFoundationApi::class)
@Composable
fun ImageViewer(list: List<Item>, start: Int, loader: ImageLoader, onClose: () -> Unit) {
    val pager = rememberPagerState(initialPage = start) { list.size }
    BackHandler(onBack = onClose)
    Box(Modifier.fillMaxSize().background(Color.Black)) {
        HorizontalPager(state = pager, modifier = Modifier.fillMaxSize()) { p ->
            AsyncImage(
                model = list[p].file, imageLoader = loader, contentDescription = null,
                contentScale = ContentScale.Fit, modifier = Modifier.fillMaxSize()
            )
        }
        IconButton(onClick = onClose, modifier = Modifier.align(Alignment.TopStart).padding(8.dp)) {
            Icon(Icons.AutoMirrored.Filled.ArrowBack, "Back", tint = Color.White)
        }
        Text(
            "${pager.currentPage + 1}/${list.size} • ${list[pager.currentPage].name}",
            color = Color.White, fontSize = 12.sp, maxLines = 1,
            modifier = Modifier.align(Alignment.BottomCenter).padding(12.dp)
        )
    }
}
