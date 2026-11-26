package com.fu.wa

import android.annotation.SuppressLint
import android.os.Bundle
import android.webkit.CookieManager
import android.webkit.WebView
import android.webkit.WebViewClient
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.compose.ui.viewinterop.AndroidView
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.json.JSONArray

private const val COURSE_JS =
    """(function(){var s={};document.querySelectorAll('a[href*="course/view.php?id="]').forEach(function(a){var m=a.href.match(/[?&]id=(\d+)/);if(m&&m[1]!=='1')s[m[1]]=1;});return Object.keys(s);})()"""

class LoginActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContent { MaterialTheme(colorScheme = fuColors()) { LoginScreen { finish() } } }
    }
}

@SuppressLint("SetJavaScriptEnabled")
@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun LoginScreen(onClose: () -> Unit) {
    val ctx = LocalContext.current
    val scope = rememberCoroutineScope()
    var status by remember {
        mutableStateOf("Apni AIOU (Microsoft) ID se login karein. Login ke baad courses khud sync ho jayenge.")
    }
    var busy by remember { mutableStateOf(false) }
    var done by remember { mutableStateOf(false) }
    var web by remember { mutableStateOf<WebView?>(null) }
    var autoTried by remember { mutableStateOf(false) }

    fun doSync() {
        val w = web ?: return
        if (busy) return
        busy = true
        status = "Courses dhoond raha hoon…"
        w.evaluateJavascript(COURSE_JS) { res ->
            val ids = runCatching {
                val a = JSONArray(res ?: "[]")
                (0 until a.length()).map { a.getString(it) }
            }.getOrDefault(emptyList()).filter { it.isNotBlank() && it != "1" }.distinct()
            if (ids.isEmpty()) {
                busy = false
                status = "Courses nahi mile. Login kar ke 'My courses' page khulne dein, phir 'Sync courses' dabayein."
            } else {
                val cm = CookieManager.getInstance()
                cm.flush()
                Store.setCookie(ctx, cm.getCookie(Lms.BASE) ?: "")
                Store.setUa(ctx, w.settings.userAgentString)
                Store.setCourses(ctx, ids)
                status = "${ids.size} courses mile — workshops parh raha hoon…"
                scope.launch {
                    val msg = withContext(Dispatchers.IO) {
                        runCatching { Lms.syncCourses(ctx, ids, notifyNew = false) }.getOrElse { "Error: ${it.message}" }
                    }
                    busy = false
                    status = msg
                    done = msg.startsWith("✔")
                }
            }
        }
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = { Text("AIOU Login & Sync", fontWeight = FontWeight.Bold) },
                colors = TopAppBarDefaults.topAppBarColors(containerColor = Navy, titleContentColor = Color.White),
                actions = { TextButton(onClick = onClose) { Text("Wapas", color = Color.White) } }
            )
        }
    ) { pad ->
        Column(Modifier.padding(pad).fillMaxSize()) {
            Column(Modifier.fillMaxWidth().background(Color(0xFFD7F3F1)).padding(12.dp)) {
                Text(status, fontSize = 14.sp, color = Navy)
                Spacer(Modifier.height(8.dp))
                Row(horizontalArrangement = Arrangement.spacedBy(8.dp), verticalAlignment = Alignment.CenterVertically) {
                    Button(onClick = { doSync() }, enabled = !busy,
                        colors = ButtonDefaults.buttonColors(containerColor = Teal)) {
                        Text(if (busy) "Sync ho raha hai…" else "Sync courses")
                    }
                    if (done) OutlinedButton(onClick = onClose) { Text("Done ✔") }
                }
            }
            AndroidView(
                factory = { c ->
                    WebView(c).apply {
                        settings.javaScriptEnabled = true
                        settings.domStorageEnabled = true
                        CookieManager.getInstance().setAcceptCookie(true)
                        CookieManager.getInstance().setAcceptThirdPartyCookies(this, true)
                        webViewClient = object : WebViewClient() {
                            override fun onPageFinished(view: WebView?, url: String?) {
                                if (!autoTried && url != null && url.startsWith(Lms.BASE + "/my")) {
                                    autoTried = true
                                    view?.postDelayed({ doSync() }, 3500)
                                }
                            }
                        }
                        loadUrl(Lms.BASE + "/my/courses.php")
                        web = this
                    }
                },
                modifier = Modifier.fillMaxWidth().weight(1f)
            )
        }
    }
}
