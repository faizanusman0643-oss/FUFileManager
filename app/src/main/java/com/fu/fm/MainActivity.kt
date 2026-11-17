package com.fu.fm

import android.Manifest
import android.app.Application
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.provider.Settings
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.viewModels
import androidx.compose.foundation.layout.*
import androidx.compose.material3.Button
import androidx.compose.material3.Text
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.unit.dp
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import coil.ImageLoader
import coil.decode.VideoFrameDecoder
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.io.File

class VM(app: Application) : AndroidViewModel(app) {
    var items by mutableStateOf<List<Item>>(emptyList())
    var loading by mutableStateOf(false)
    var busy by mutableStateOf(false)
    var version by mutableIntStateOf(0)
    var filter by mutableStateOf<Cat?>(null)
    var message by mutableStateOf<String?>(null)
    var vaultOpen by mutableStateOf(false)
    var dupes by mutableStateOf<List<List<Item>>?>(null)
    val selected = mutableStateListOf<String>()

    private val ctx get() = getApplication<Application>()

    fun refresh() {
        viewModelScope.launch {
            loading = true
            items = withContext(Dispatchers.IO) { Fm.scan().sortedByDescending { it.mtime } }
            version++
            loading = false
        }
    }

    fun op(work: () -> String) {
        viewModelScope.launch {
            busy = true
            val r = withContext(Dispatchers.IO) {
                try { work() } catch (e: Exception) { "Error: ${e.message}" }
            }
            busy = false
            message = r
            selected.clear()
            refresh()
        }
    }

    fun toggle(path: String) { if (!selected.remove(path)) selected.add(path) }

    private fun rescan(m: List<Moved>) =
        Fm.rescan(ctx, m.flatMap { listOf(it.from.absolutePath, it.to.absolutePath) })

    fun delete(paths: List<String>) = op {
        val done = Fm.deleteAll(paths.map { File(it) })
        Fm.rescan(ctx, done)
        "${done.size} delete hue"
    }

    fun move(paths: List<String>, dest: File, copy: Boolean) = op {
        val r = Fm.moveAll(paths.map { File(it) }, dest, copy)
        rescan(r)
        "${r.size} ${if (copy) "copy" else "move"} hue"
    }

    fun hide(paths: List<String>) = op {
        if (!Pin.isSet(ctx)) "Pehle Vault tab mai PIN set karen"
        else {
            val n = Vault.hide(paths.map { File(it) })
            Fm.rescan(ctx, paths)
            "$n vault mai chhupa diye"
        }
    }

    fun restore(names: List<String>) = op { "${Vault.restore(names)} wapas aa gaye" }

    fun sort(dir: File, mode: SortMode) = op {
        val r = Fm.sortFolder(dir, mode)
        rescan(r)
        "${r.size} files sort hui"
    }

    fun consolidate(src: File, dst: File) = op {
        val r = Fm.consolidate(src, dst)
        rescan(r)
        "${r.size} files ek jagah aa gayi"
    }

    fun emptyDirs(dir: File) = op { "${Fm.deleteEmptyDirs(dir)} empty folders delete hue" }

    fun clean() = op {
        val (b, n) = Fm.cleanJunk()
        "${fmtSize(b)} saaf, $n empty folders delete"
    }

    fun findDupes() {
        viewModelScope.launch {
            busy = true
            val r = withContext(Dispatchers.IO) { Fm.findDuplicates(items) }
            busy = false
            if (r.isEmpty()) message = "Koi duplicate nahi mila" else dupes = r
        }
    }
}

class MainActivity : ComponentActivity() {
    private var granted by mutableStateOf(false)
    private var external by mutableStateOf<Uri?>(null)
    private val vm: VM by viewModels()

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        handleIntent(intent)
        val loader = ImageLoader.Builder(this)
            .components { add(VideoFrameDecoder.Factory()) }
            .crossfade(true)
            .build()
        setContent {
            FuTheme {
                if (granted) {
                    LaunchedEffect(Unit) { vm.refresh() }
                    App(vm, loader, external) { external = null }
                } else {
                    PermissionScreen { requestAccess() }
                }
            }
        }
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        handleIntent(intent)
    }

    private fun handleIntent(i: Intent?) {
        if (i?.action == Intent.ACTION_VIEW && i.data != null) external = i.data
    }

    override fun onResume() {
        super.onResume()
        granted = hasAccess()
    }

    private fun hasAccess(): Boolean =
        if (Build.VERSION.SDK_INT >= 30) Environment.isExternalStorageManager()
        else ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE) ==
            PackageManager.PERMISSION_GRANTED

    private fun requestAccess() {
        if (Build.VERSION.SDK_INT >= 30) {
            try {
                startActivity(Intent(Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION, Uri.parse("package:$packageName")))
            } catch (e: Exception) {
                startActivity(Intent(Settings.ACTION_MANAGE_ALL_FILES_ACCESS_PERMISSION))
            }
        } else {
            ActivityCompat.requestPermissions(
                this,
                arrayOf(Manifest.permission.READ_EXTERNAL_STORAGE, Manifest.permission.WRITE_EXTERNAL_STORAGE),
                1
            )
        }
    }

    override fun onRequestPermissionsResult(requestCode: Int, permissions: Array<out String>, grantResults: IntArray) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        granted = hasAccess()
    }
}

@Composable
fun PermissionScreen(onGrant: () -> Unit) {
    Column(
        Modifier.fillMaxSize().padding(24.dp),
        verticalArrangement = Arrangement.Center,
        horizontalAlignment = Alignment.CenterHorizontally
    ) {
        Text("FU File Manager", style = androidx.compose.material3.MaterialTheme.typography.headlineSmall)
        Spacer(Modifier.height(12.dp))
        Text(
            "Sab files dikhane, move aur delete karne ke liye 'All files access' chahiye. Neeche button dabao aur is app ke liye allow kar do.",
            textAlign = TextAlign.Center
        )
        Spacer(Modifier.height(20.dp))
        Button(onClick = onGrant) { Text("Allow access") }
    }
}
