package com.fu.xplayer

import android.annotation.SuppressLint
import android.graphics.Bitmap
import android.net.Uri
import android.webkit.*
import androidx.activity.compose.BackHandler
import androidx.compose.foundation.*
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.compose.ui.viewinterop.AndroidView
import java.io.ByteArrayInputStream

data class Found(val url: String, val page: String, val kind: String)

private val mediaRx = Regex("\\.(mp4|m4v|webm|mkv|mov|avi|flv|3gp|wmv|mp3|m4a|aac|wav|ogg|opus|flac|m3u8|mpd)(\\?|#|$)", RegexOption.IGNORE_CASE)
private val adHosts = listOf(
    "doubleclick.net", "googlesyndication.com", "googleadservices.com", "adservice.google", "popads.net", "popcash.net",
    "propellerads", "adsterra", "exoclick.com", "juicyads.com", "trafficjunky", "adnxs.com", "taboola.com", "outbrain.com",
    "mgid.com", "onclickads", "hilltopads", "a-ads.com", "yllix.com", "clickadu", "ad.plus", "pubmatic.com", "criteo.com"
)

object BrowserState {
    var pending by mutableStateOf<String?>(null)
}

@SuppressLint("SetJavaScriptEnabled")
@Composable
fun BrowserScreen(onMin: () -> Unit, onClose: () -> Unit, onDownload: (String, String?, Map<String, String>) -> Unit, onPlay: (PItem) -> Unit) {
    var web by remember { mutableStateOf<WebView?>(null) }
    var address by remember { mutableStateOf(Prefs.homePage) }
    var title by remember { mutableStateOf("FU Browser") }
    var progress by remember { mutableIntStateOf(0) }
    var canBack by remember { mutableStateOf(false) }
    val found = remember { mutableStateListOf<Found>() }
    var showFound by remember { mutableStateOf(false) }
    var menu by remember { mutableStateOf(false) }
    var blocked by remember { mutableIntStateOf(0) }

    fun go(text: String) {
        val t = text.trim()
        val url = when {
            t.startsWith("http://") || t.startsWith("https://") -> t
            t.contains('.') && !t.contains(' ') -> "https://$t"
            else -> "https://www.google.com/search?q=" + Uri.encode(t)
        }
        web?.loadUrl(url)
    }

    LaunchedEffect(BrowserState.pending, web) {
        val p = BrowserState.pending
        if (p != null && web != null) { go(p); BrowserState.pending = null }
    }

    BackHandler(enabled = canBack) { web?.goBack() }

    WinWindow(title = "$title - FU Browser", icon = Icons.Filled.Public, iconTint = Color(0xFF1E88E5), onMin = onMin, onClose = onClose) {
        Row(Modifier.fillMaxWidth().background(W7.cmd).padding(4.dp), verticalAlignment = Alignment.CenterVertically) {
            IconButton({ web?.goBack() }, Modifier.size(36.dp)) { Icon(Icons.Filled.ArrowBack, "Back", tint = if (canBack) W7.link else Color.LightGray) }
            IconButton({ web?.goForward() }, Modifier.size(36.dp)) { Icon(Icons.Filled.ArrowForward, "Forward", tint = W7.link) }
            WinField(address, { address = it }, Modifier.weight(1f), "Search ya URL likhen", onDone = { go(address) })
            IconButton({ web?.reload() }, Modifier.size(36.dp)) { Icon(Icons.Filled.Refresh, "Refresh", tint = W7.link) }
            Box {
                IconButton({ menu = true }, Modifier.size(36.dp)) { Icon(Icons.Filled.MoreVert, "Menu", tint = W7.link) }
                DropdownMenu(menu, { menu = false }) {
                    DropdownMenuItem({ Text("Home") }, { menu = false; go(Prefs.homePage) }, leadingIcon = { Icon(Icons.Filled.Home, null) })
                    DropdownMenuItem({ Text("Is page ko home banao") }, { menu = false; web?.url?.let { Prefs.homePage = it; Prefs.save() } })
                    DropdownMenuItem({ Text(if (Prefs.desktopMode) "✓ Desktop site" else "Desktop site") }, {
                        menu = false; Prefs.desktopMode = !Prefs.desktopMode; Prefs.save()
                        web?.settings?.userAgentString = if (Prefs.desktopMode) "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36" else null
                        web?.reload()
                    })
                    DropdownMenuItem({ Text(if (Prefs.adBlock) "✓ Ad blocker ($blocked blocked)" else "Ad blocker") }, { menu = false; Prefs.adBlock = !Prefs.adBlock; Prefs.save() })
                    DropdownMenuItem({ Text("Is page ka link download karo") }, { menu = false; web?.url?.let { onDownload(it, null, headersFor(it, it)) } })
                    DropdownMenuItem({ Text("History / cookies saaf karo") }, {
                        menu = false; web?.clearHistory(); web?.clearCache(true); CookieManager.getInstance().removeAllCookies(null)
                    })
                }
            }
        }
        if (progress in 1..99) LinearProgressIndicator({ progress / 100f }, Modifier.fillMaxWidth().height(3.dp), color = W7.teal)
        Box(Modifier.weight(1f).fillMaxWidth()) {
            AndroidView(factory = { c ->
                WebView(c).apply {
                    settings.javaScriptEnabled = true
                    settings.domStorageEnabled = true
                    settings.loadWithOverviewMode = true
                    settings.useWideViewPort = true
                    settings.builtInZoomControls = true
                    settings.displayZoomControls = false
                    settings.mediaPlaybackRequiresUserGesture = false
                    settings.setSupportMultipleWindows(false)
                    if (Prefs.desktopMode) settings.userAgentString = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
                    CookieManager.getInstance().setAcceptThirdPartyCookies(this, true)
                    webChromeClient = object : WebChromeClient() {
                        override fun onProgressChanged(view: WebView?, newProgress: Int) { progress = newProgress }
                        override fun onReceivedTitle(view: WebView?, t: String?) { title = t ?: "" }
                    }
                    webViewClient = object : WebViewClient() {
                        override fun onPageStarted(view: WebView?, url: String?, favicon: Bitmap?) {
                            address = url ?: ""; found.clear(); canBack = view?.canGoBack() == true
                        }
                        override fun onPageFinished(view: WebView?, url: String?) { canBack = view?.canGoBack() == true }
                        override fun shouldOverrideUrlLoading(view: WebView?, request: WebResourceRequest?): Boolean {
                            val u = request?.url?.toString() ?: return false
                            if (!u.startsWith("http")) return true // block intent://, market:// popups
                            return false
                        }
                        override fun shouldInterceptRequest(view: WebView?, request: WebResourceRequest?): WebResourceResponse? {
                            val u = request?.url?.toString() ?: return null
                            val host = request.url.host ?: ""
                            if (Prefs.adBlock && adHosts.any { host.contains(it) }) {
                                view?.post { blocked++ }
                                return WebResourceResponse("text/plain", "utf-8", ByteArrayInputStream(ByteArray(0)))
                            }
                            val m = mediaRx.find(u.substringBefore('#'))
                            if (m != null) {
                                val k = m.groupValues[1].lowercase()
                                view?.post {
                                    if (found.none { it.url == u }) found.add(Found(u, view.url ?: u, k))
                                }
                            }
                            return null
                        }
                    }
                    setDownloadListener { url, ua, cd, mime, _ ->
                        val name = Dl.guessName(url, cd, mime)
                        onDownload(url, name, headersFor(url, this.url ?: url, ua))
                    }
                    setOnLongClickListener {
                        val h = hitTestResult
                        val u = h.extra
                        if (u != null && (h.type == WebView.HitTestResult.SRC_ANCHOR_TYPE || h.type == WebView.HitTestResult.IMAGE_TYPE ||
                                h.type == WebView.HitTestResult.SRC_IMAGE_ANCHOR_TYPE) && u.startsWith("http")
                        ) { onDownload(u, null, headersFor(u, this.url ?: u)); true } else false
                    }
                    loadUrl(BrowserState.pending ?: Prefs.homePage)
                    BrowserState.pending = null
                    web = this
                }
            }, modifier = Modifier.fillMaxSize())

            if (found.isNotEmpty()) {
                Row(
                    Modifier.align(Alignment.BottomEnd).padding(14.dp).clip(RoundedCornerShape(24.dp)).background(W7.green)
                        .border(1.dp, Color(0xFF1B7A1B), RoundedCornerShape(24.dp)).clickable { showFound = true }
                        .padding(horizontal = 16.dp, vertical = 10.dp),
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Icon(Icons.Filled.Download, null, tint = Color.White)
                    Text(" Download (${found.size})", color = Color.White, fontSize = 15.sp)
                }
            }
        }
    }

    DisposableEffect(Unit) { onDispose { web?.destroy() } }

    if (showFound) {
        WinDialog("Is page pe media mili (${found.size})", { showFound = false }, Icons.Filled.Download, W7.teal, buttons = {
            WinButton("Close") { showFound = false }
        }) {
            LazyColumn(Modifier.heightIn(max = 440.dp)) {
                items(found.toList()) { f ->
                    Column(Modifier.fillMaxWidth().padding(vertical = 6.dp)) {
                        Text(Dl.guessName(f.url), maxLines = 1, overflow = TextOverflow.Ellipsis, fontSize = 14.sp, color = W7.text)
                        Text("${f.kind.uppercase()} • ${Uri.parse(f.url).host}", fontSize = 11.sp, color = W7.gray)
                        Row(horizontalArrangement = Arrangement.spacedBy(6.dp), modifier = Modifier.padding(top = 4.dp)) {
                            WinButton("Download", default = true) { showFound = false; onDownload(f.url, null, headersFor(f.url, f.page)) }
                            WinButton("Play") { showFound = false; onPlay(PItem(Uri.parse(f.url), Dl.guessName(f.url), f.kind in setOf("mp3", "m4a", "aac", "wav", "ogg", "opus", "flac"))) }
                        }
                        HorizontalDivider(Modifier.padding(top = 6.dp), color = Color(0xFFE0E0E0))
                    }
                }
            }
        }
    }
}

fun headersFor(url: String, referer: String, ua: String? = null): Map<String, String> {
    val m = HashMap<String, String>()
    m["Referer"] = referer
    m["User-Agent"] = ua ?: Dl.UA
    CookieManager.getInstance().getCookie(url)?.let { m["Cookie"] = it }
    return m
}
