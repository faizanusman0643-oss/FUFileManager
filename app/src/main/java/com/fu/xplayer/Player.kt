package com.fu.xplayer

import android.app.Activity
import android.content.pm.ActivityInfo
import android.net.Uri
import android.view.ViewGroup
import androidx.annotation.OptIn
import androidx.compose.foundation.*
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.itemsIndexed
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Brush
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.compose.ui.viewinterop.AndroidView
import androidx.media3.common.MediaItem
import androidx.media3.common.MediaMetadata
import androidx.media3.common.MimeTypes
import androidx.media3.common.PlaybackException
import androidx.media3.common.Player
import androidx.media3.common.util.UnstableApi
import androidx.media3.ui.AspectRatioFrameLayout
import androidx.media3.ui.PlayerView
import kotlinx.coroutines.delay
import java.io.File

data class PItem(val uri: Uri, val title: String, val audio: Boolean)

fun File.toPItem() = PItem(Uri.fromFile(this), name, Fs.kind(this) == Kind.AUDIO)

/** Current queue shared app-wide (taskbar shows "Now playing"). */
object Queue {
    var items by mutableStateOf<List<PItem>>(emptyList())
    var start by mutableIntStateOf(0)
    var version by mutableIntStateOf(0)
    fun play(list: List<PItem>, index: Int) { items = list; start = index.coerceIn(0, (list.size - 1).coerceAtLeast(0)); version++ }
}

private fun mediaItem(p: PItem): MediaItem {
    val b = MediaItem.Builder().setUri(p.uri).setMediaId(p.uri.toString())
        .setMediaMetadata(MediaMetadata.Builder().setTitle(p.title).build())
    val path = p.uri.toString().substringBefore('?').lowercase()
    if (path.endsWith(".m3u8")) b.setMimeType(MimeTypes.APPLICATION_M3U8)
    if (path.endsWith(".mpd")) b.setMimeType(MimeTypes.APPLICATION_MPD)
    return b.build()
}

@OptIn(UnstableApi::class)
@Composable
fun PlayerScreen(onClose: () -> Unit, onMin: () -> Unit) {
    val ctx = LocalContext.current
    val act = ctx as? Activity
    val player = remember { Players.get(ctx) }
    var title by remember { mutableStateOf("") }
    var isAudio by remember { mutableStateOf(false) }
    var speed by remember { mutableFloatStateOf(1f) }
    var repeat by remember { mutableIntStateOf(player.repeatMode) }
    var shuffle by remember { mutableStateOf(player.shuffleModeEnabled) }
    var resize by remember { mutableIntStateOf(AspectRatioFrameLayout.RESIZE_MODE_FIT) }
    var showList by remember { mutableStateOf(false) }
    var speedMenu by remember { mutableStateOf(false) }
    var landscape by remember { mutableStateOf(false) }
    var err by remember { mutableStateOf<String?>(null) }
    var index by remember { mutableIntStateOf(0) }
    var timerMin by remember { mutableIntStateOf(0) }
    var timerMenu by remember { mutableStateOf(false) }

    // load queue
    LaunchedEffect(Queue.version) {
        if (Queue.items.isEmpty()) return@LaunchedEffect
        val same = player.mediaItemCount == Queue.items.size &&
            (0 until player.mediaItemCount).all { player.getMediaItemAt(it).mediaId == Queue.items[it].uri.toString() }
        if (!same || player.currentMediaItemIndex != Queue.start) {
            player.setMediaItems(Queue.items.map { mediaItem(it) }, Queue.start, Prefs.getPos(Queue.items[Queue.start].uri.toString()).let { if (it > 5000) it else 0L })
            player.prepare()
        }
        player.play()
    }

    DisposableEffect(player) {
        fun sync() {
            index = player.currentMediaItemIndex
            val it = Queue.items.getOrNull(index)
            title = player.currentMediaItem?.mediaMetadata?.title?.toString() ?: it?.title ?: ""
            isAudio = it?.audio ?: (player.videoSize.width == 0)
        }
        val l = object : Player.Listener {
            override fun onMediaItemTransition(mediaItem: MediaItem?, reason: Int) { sync(); err = null }
            override fun onEvents(p: Player, events: Player.Events) { sync() }
            override fun onPlayerError(error: PlaybackException) { err = "Play nahi ho saki: ${error.errorCodeName}" }
        }
        player.addListener(l); sync()
        onDispose { player.removeListener(l) }
    }

    // remember resume position
    LaunchedEffect(Unit) {
        while (true) {
            delay(3000)
            player.currentMediaItem?.mediaId?.let { id ->
                val pos = player.currentPosition; val dur = player.duration
                Prefs.setPos(id, if (dur > 0 && pos > dur - 5000) 0 else pos)
            }
        }
    }

    // sleep timer
    LaunchedEffect(timerMin) {
        if (timerMin > 0) { delay(timerMin * 60_000L); player.pause(); timerMin = 0 }
    }

    DisposableEffect(landscape) {
        act?.requestedOrientation = if (landscape) ActivityInfo.SCREEN_ORIENTATION_SENSOR_LANDSCAPE else ActivityInfo.SCREEN_ORIENTATION_UNSPECIFIED
        onDispose { }
    }
    DisposableEffect(Unit) { onDispose { act?.requestedOrientation = ActivityInfo.SCREEN_ORIENTATION_UNSPECIFIED } }

    WinWindow(
        title = if (title.isNotEmpty()) "$title - FU Media Player" else "FU Media Player",
        icon = Icons.Filled.PlayCircle, iconTint = Color(0xFFE88A1A),
        onMin = onMin, onClose = onClose
    ) {
        Box(Modifier.weight(1f).fillMaxWidth().background(Color.Black)) {
            if (isAudio) {
                Column(
                    Modifier.fillMaxSize().background(Brush.verticalGradient(listOf(Color(0xFF0B2545), Color(0xFF13A8A0), Color(0xFF0B2545)))),
                    horizontalAlignment = Alignment.CenterHorizontally, verticalArrangement = Arrangement.Center
                ) {
                    Icon(Icons.Filled.MusicNote, null, tint = Color.White, modifier = Modifier.size(120.dp))
                    Text(title, color = Color.White, fontSize = 18.sp, fontWeight = FontWeight.Bold, modifier = Modifier.padding(16.dp), maxLines = 2)
                }
            }
            AndroidView(
                factory = { c ->
                    PlayerView(c).apply {
                        layoutParams = ViewGroup.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, ViewGroup.LayoutParams.MATCH_PARENT)
                        this.player = player
                        setShowNextButton(true); setShowPreviousButton(true)
                        setShowShuffleButton(true)
                        setShowSubtitleButton(true)
                        setShutterBackgroundColor(android.graphics.Color.TRANSPARENT)
                        setBackgroundColor(android.graphics.Color.TRANSPARENT)
                        controllerShowTimeoutMs = 3500
                        setFullscreenButtonClickListener { full -> landscape = full }
                        keepScreenOn = true
                    }
                },
                update = { v -> v.resizeMode = resize; v.player = player },
                modifier = Modifier.fillMaxSize()
            )
            err?.let {
                Text(it, color = Color.White, modifier = Modifier.align(Alignment.TopCenter).background(Color(0xCCB00020)).padding(8.dp))
            }
        }
        // WMP-style control strip
        Row(
            Modifier.fillMaxWidth().background(Brush.verticalGradient(listOf(Color(0xFF2C3E50), Color(0xFF0E1620))))
                .horizontalScroll(rememberScrollState()).padding(horizontal = 4.dp, vertical = 2.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            val tint = Color(0xFFBFE6FF)
            IconButton({ player.seekToPreviousMediaItem() }) { Icon(Icons.Filled.SkipPrevious, "Prev", tint = tint) }
            IconButton({ player.seekBack() }) { Icon(Icons.Filled.Replay10, "-10s", tint = tint) }
            IconButton({ if (player.isPlaying) player.pause() else player.play() }) { Icon(Icons.Filled.PlayArrow, "Play/Pause", tint = Color.White) }
            IconButton({ player.seekForward() }) { Icon(Icons.Filled.Forward10, "+10s", tint = tint) }
            IconButton({ player.seekToNextMediaItem() }) { Icon(Icons.Filled.SkipNext, "Next", tint = tint) }
            IconButton({
                repeat = when (repeat) { Player.REPEAT_MODE_OFF -> Player.REPEAT_MODE_ALL; Player.REPEAT_MODE_ALL -> Player.REPEAT_MODE_ONE; else -> Player.REPEAT_MODE_OFF }
                player.repeatMode = repeat
            }) {
                Icon(if (repeat == Player.REPEAT_MODE_ONE) Icons.Filled.RepeatOne else Icons.Filled.Repeat, "Repeat",
                    tint = if (repeat == Player.REPEAT_MODE_OFF) Color.Gray else W7.teal)
            }
            IconButton({ shuffle = !shuffle; player.shuffleModeEnabled = shuffle }) {
                Icon(Icons.Filled.Shuffle, "Shuffle", tint = if (shuffle) W7.teal else Color.Gray)
            }
            Box {
                TextButton({ speedMenu = true }) { Text("${speed}x", color = tint) }
                DropdownMenu(speedMenu, { speedMenu = false }) {
                    listOf(0.25f, 0.5f, 0.75f, 1f, 1.25f, 1.5f, 1.75f, 2f, 3f).forEach { s ->
                        DropdownMenuItem({ Text("${s}x") }, { speed = s; player.setPlaybackSpeed(s); speedMenu = false })
                    }
                }
            }
            IconButton({
                resize = when (resize) {
                    AspectRatioFrameLayout.RESIZE_MODE_FIT -> AspectRatioFrameLayout.RESIZE_MODE_ZOOM
                    AspectRatioFrameLayout.RESIZE_MODE_ZOOM -> AspectRatioFrameLayout.RESIZE_MODE_FILL
                    else -> AspectRatioFrameLayout.RESIZE_MODE_FIT
                }
            }) { Icon(Icons.Filled.AspectRatio, "Fit/Zoom/Stretch", tint = tint) }
            IconButton({ landscape = !landscape }) { Icon(Icons.Filled.ScreenRotation, "Rotate", tint = tint) }
            Box {
                IconButton({ timerMenu = true }) { Icon(Icons.Filled.Timer, "Sleep timer", tint = if (timerMin > 0) W7.teal else tint) }
                DropdownMenu(timerMenu, { timerMenu = false }) {
                    listOf(0, 10, 20, 30, 45, 60, 90).forEach { m ->
                        DropdownMenuItem({ Text(if (m == 0) "Timer band" else "$m minute baad band") }, { timerMin = m; timerMenu = false })
                    }
                }
            }
            IconButton({ showList = true }) { Icon(Icons.Filled.QueueMusic, "Playlist", tint = tint) }
        }
    }

    if (showList) {
        WinDialog("Now Playing list (${Queue.items.size})", { showList = false }, Icons.Filled.QueueMusic, buttons = {
            WinButton("Close", default = true) { showList = false }
        }) {
            LazyColumn(Modifier.heightIn(max = 420.dp)) {
                itemsIndexed(Queue.items) { i, it ->
                    Text(
                        "${i + 1}. ${it.title}", maxLines = 1, overflow = TextOverflow.Ellipsis,
                        color = if (i == index) W7.link else W7.text,
                        fontWeight = if (i == index) FontWeight.Bold else FontWeight.Normal,
                        modifier = Modifier.fillMaxWidth().background(if (i == index) W7.sel else Color.Transparent)
                            .clickable { player.seekTo(i, 0); player.play(); showList = false }.padding(8.dp)
                    )
                }
            }
        }
    }
}
