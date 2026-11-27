package com.fu.xplayer

import androidx.compose.foundation.*
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.shape.CircleShape
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.text.BasicTextField
import androidx.compose.foundation.text.KeyboardActions
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.draw.shadow
import androidx.compose.ui.layout.layout
import androidx.compose.ui.graphics.Brush
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.SolidColor
import androidx.compose.ui.graphics.vector.ImageVector
import androidx.compose.ui.text.TextStyle
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.ImeAction
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.compose.ui.window.Dialog
import kotlinx.coroutines.delay
import java.io.File
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

/** Windows 7 "Aero" look: colors, frames, buttons, taskbar. */
object W7 {
    val desk = Brush.linearGradient(
        listOf(Color(0xFF062A55), Color(0xFF0F4F94), Color(0xFF1E78C8), Color(0xFF3A9BE0), Color(0xFF0D4A8A))
    )
    val aero = Brush.verticalGradient(
        listOf(Color(0xFFE3EFFB), Color(0xFFC4DAF1), Color(0xFFA9C7E8), Color(0xFFBCD4EE))
    )
    val aeroInactive = Brush.verticalGradient(listOf(Color(0xFFEEF2F6), Color(0xFFD6DEE7)))
    val cmd = Brush.verticalGradient(listOf(Color(0xFFFAFCFE), Color(0xFFE9F0F8), Color(0xFFDCE6F3)))
    val taskbar = Brush.verticalGradient(
        listOf(Color(0xE8597A99), Color(0xF0233A52), Color(0xFA101C28))
    )
    val orb = Brush.radialGradient(listOf(Color(0xFFB8F0FF), Color(0xFF2EA3E8), Color(0xFF0B4A8A), Color(0xFF062A55)))
    val btn = Brush.verticalGradient(listOf(Color(0xFFF6F6F6), Color(0xFFEBEBEB), Color(0xFFDDDDDD), Color(0xFFCFCFCF)))
    val btnBlue = Brush.verticalGradient(listOf(Color(0xFFEAF6FD), Color(0xFFD9F0FC), Color(0xFFBEE6FD), Color(0xFFA7D9F5)))
    val close = Brush.verticalGradient(listOf(Color(0xFFE9A99B), Color(0xFFD5583F), Color(0xFFC13A1F), Color(0xFFD86F55)))
    val green = Brush.verticalGradient(listOf(Color(0xFF8BE38B), Color(0xFF2FB52F), Color(0xFF1E9B1E)))
    val startRight = Brush.verticalGradient(listOf(Color(0xE0284D73), Color(0xF0173450)))
    val sel = Color(0xFFCCE8FF)
    val selBorder = Color(0xFF7DA2CE)
    val text = Color(0xFF1E1E1E)
    val gray = Color(0xFF6D6D6D)
    val link = Color(0xFF1E5BA8)
    val border = Color(0xFF8E9FB4)
    val navy = Color(0xFF0B2545)
    val teal = Color(0xFF13A8A0)
    val navBg = Color(0xFFF1F5FB)
    val folder = Color(0xFFE8B931)
}

@Composable
fun W7Theme(content: @Composable () -> Unit) {
    MaterialTheme(
        colorScheme = lightColorScheme(
            primary = Color(0xFF1E5BA8),
            secondary = W7.teal,
            surface = Color(0xFFF5F5F5),
            surfaceContainer = Color(0xFFF5F5F5),
            background = Color.White
        ),
        content = content
    )
}

/** A Win7 window: aero title bar + white client area. */
@Composable
fun WinWindow(
    title: String,
    icon: ImageVector,
    iconTint: Color = W7.link,
    modifier: Modifier = Modifier,
    onMin: (() -> Unit)? = null,
    onClose: (() -> Unit)? = null,
    content: @Composable ColumnScope.() -> Unit
) {
    Column(
        modifier
            .fillMaxSize()
            .background(W7.aero)
            .border(1.dp, Color(0xFF5A7FA8))
            .padding(horizontal = 3.dp)
    ) {
        Row(Modifier.fillMaxWidth().height(32.dp), verticalAlignment = Alignment.CenterVertically) {
            Icon(icon, null, tint = iconTint, modifier = Modifier.padding(start = 4.dp).size(18.dp))
            Text(
                title, Modifier.weight(1f).padding(start = 6.dp), color = Color(0xFF1B2B3D),
                fontSize = 13.sp, maxLines = 1, overflow = TextOverflow.Ellipsis
            )
            if (onMin != null) Caption("—", W7.btnBlue, Color(0xFF26364A), onMin)
            if (onClose != null) Caption("✕", W7.close, Color.White, onClose, wide = true)
        }
        Column(
            Modifier.fillMaxSize().padding(bottom = 3.dp).background(Color.White).border(1.dp, Color(0xFF8AA4C2)),
            content = content
        )
    }
}

@Composable
private fun Caption(t: String, b: Brush, c: Color, onClick: () -> Unit, wide: Boolean = false) {
    Box(
        Modifier.padding(start = 2.dp).size(if (wide) 44.dp else 28.dp, 20.dp)
            .clip(RoundedCornerShape(bottomStart = 4.dp, bottomEnd = 4.dp))
            .background(b).border(1.dp, Color(0xFF4E6B8C), RoundedCornerShape(bottomStart = 4.dp, bottomEnd = 4.dp))
            .clickable(onClick = onClick),
        contentAlignment = Alignment.Center
    ) { Text(t, color = c, fontSize = 11.sp, fontWeight = FontWeight.Bold) }
}

@Composable
fun WinButton(text: String, modifier: Modifier = Modifier, enabled: Boolean = true, default: Boolean = false, onClick: () -> Unit) {
    Box(
        modifier.heightIn(min = 34.dp).widthIn(min = 76.dp)
            .clip(RoundedCornerShape(3.dp))
            .background(if (default) W7.btnBlue else W7.btn)
            .border(1.dp, if (default) Color(0xFF3C7FB1) else Color(0xFF707070), RoundedCornerShape(3.dp))
            .clickable(enabled = enabled, onClick = onClick)
            .padding(horizontal = 12.dp, vertical = 6.dp),
        contentAlignment = Alignment.Center
    ) { Text(text, color = if (enabled) W7.text else Color(0xFFA0A0A0), fontSize = 14.sp) }
}

/** Flat command-bar item (Organize ▾, New folder ...). */
@Composable
fun CmdItem(text: String?, icon: ImageVector?, tint: Color = Color(0xFF1E3A5F), onClick: () -> Unit) {
    Row(
        Modifier.clip(RoundedCornerShape(3.dp)).clickable(onClick = onClick).padding(horizontal = 7.dp, vertical = 7.dp),
        verticalAlignment = Alignment.CenterVertically
    ) {
        if (icon != null) Icon(icon, null, tint = tint, modifier = Modifier.size(20.dp))
        if (text != null) Text(text, Modifier.padding(start = if (icon != null) 4.dp else 0.dp), fontSize = 13.sp, color = W7.text)
    }
}

@Composable
fun WinField(
    value: String, onChange: (String) -> Unit, modifier: Modifier = Modifier, hint: String = "",
    onDone: (() -> Unit)? = null, singleLine: Boolean = true, password: Boolean = false
) {
    Box(
        modifier.background(Color.White).border(1.dp, Color(0xFFABADB3)).padding(horizontal = 6.dp, vertical = 8.dp)
    ) {
        if (value.isEmpty()) Text(hint, color = Color(0xFF8A8A8A), fontSize = 14.sp, fontStyle = androidx.compose.ui.text.font.FontStyle.Italic)
        BasicTextField(
            value, onChange, Modifier.fillMaxWidth(), singleLine = singleLine,
            textStyle = TextStyle(fontSize = 14.sp, color = W7.text),
            cursorBrush = SolidColor(Color.Black),
            visualTransformation = if (password) androidx.compose.ui.text.input.PasswordVisualTransformation() else androidx.compose.ui.text.input.VisualTransformation.None,
            keyboardOptions = KeyboardOptions(imeAction = if (onDone != null) ImeAction.Go else ImeAction.Default),
            keyboardActions = KeyboardActions(onGo = { onDone?.invoke() })
        )
    }
}

/** Win7 style message box dialog. */
@Composable
fun WinDialog(
    title: String,
    onDismiss: () -> Unit,
    icon: ImageVector = Icons.Filled.Info,
    iconTint: Color = W7.link,
    buttons: @Composable RowScope.() -> Unit,
    content: @Composable ColumnScope.() -> Unit
) {
    Dialog(onDismissRequest = onDismiss) {
        Column(
            Modifier.fillMaxWidth().shadow(12.dp).background(W7.aero).border(1.dp, Color(0xFF5A7FA8)).padding(horizontal = 4.dp)
        ) {
            Row(Modifier.fillMaxWidth().height(30.dp), verticalAlignment = Alignment.CenterVertically) {
                Text(title, Modifier.weight(1f).padding(start = 6.dp), fontSize = 13.sp, color = Color(0xFF1B2B3D), maxLines = 1)
                Caption("✕", W7.close, Color.White, onDismiss, wide = true)
            }
            Column(Modifier.fillMaxWidth().padding(bottom = 4.dp).background(Color.White).border(1.dp, Color(0xFF8AA4C2))) {
                Row(Modifier.padding(14.dp)) {
                    Icon(icon, null, tint = iconTint, modifier = Modifier.size(34.dp))
                    Spacer(Modifier.width(12.dp))
                    Column(Modifier.weight(1f), content = content)
                }
                Row(
                    Modifier.fillMaxWidth().background(Color(0xFFF0F0F0)).padding(10.dp),
                    horizontalArrangement = Arrangement.spacedBy(8.dp, Alignment.End),
                    verticalAlignment = Alignment.CenterVertically,
                    content = buttons
                )
            }
        }
    }
}

/** Green Win7 progress bar. */
@Composable
fun WinProgress(fraction: Float?, modifier: Modifier = Modifier, height: Int = 16) {
    Box(modifier.fillMaxWidth().height(height.dp).background(Color(0xFFE6E6E6)).border(1.dp, Color(0xFFBCBCBC))) {
        if (fraction != null) {
            Box(Modifier.fillMaxHeight().fillMaxWidth(fraction.coerceIn(0f, 1f)).padding(1.dp).background(W7.green))
        } else {
            val t = rememberInfiniteTransitionCompat()
            Box(Modifier.fillMaxHeight().fillMaxWidth(0.3f).offsetFraction(t).padding(1.dp).background(W7.green))
        }
    }
}

@Composable
private fun rememberInfiniteTransitionCompat(): Float {
    var p by remember { mutableFloatStateOf(0f) }
    LaunchedEffect(Unit) { while (true) { delay(40); p = (p + 0.02f) % 1f } }
    return p
}

private fun Modifier.offsetFraction(p: Float): Modifier = this.layout { measurable, constraints ->
    val pl = measurable.measure(constraints)
    val full = (pl.width / 0.3f)
    layout(pl.width, pl.height) { pl.placeRelative((full * 0.7f * p).toInt(), 0) }
}

// ---------------- Taskbar & Start menu ----------------

data class TaskApp(val id: Screen, val label: String, val icon: ImageVector, val tint: Color)

@Composable
fun Taskbar(apps: List<TaskApp>, front: Screen, onStart: () -> Unit, onApp: (Screen) -> Unit, onDesktop: () -> Unit) {
    var now by remember { mutableStateOf(Date()) }
    LaunchedEffect(Unit) { while (true) { now = Date(); delay(15_000) } }
    Row(
        Modifier.fillMaxWidth().height(46.dp).background(W7.taskbar),
        verticalAlignment = Alignment.CenterVertically
    ) {
        Box(
            Modifier.padding(horizontal = 6.dp).size(40.dp).shadow(6.dp, CircleShape).clip(CircleShape)
                .background(W7.orb).border(1.dp, Color(0x99FFFFFF), CircleShape).clickable(onClick = onStart),
            contentAlignment = Alignment.Center
        ) { Text("FU", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 14.sp) }
        Row(Modifier.weight(1f).horizontalScroll(rememberScrollState()), verticalAlignment = Alignment.CenterVertically) {
            apps.forEach { a ->
                val active = a.id == front
                Box(
                    Modifier.padding(horizontal = 2.dp).size(54.dp, 40.dp).clip(RoundedCornerShape(3.dp))
                        .background(if (active) Brush.verticalGradient(listOf(Color(0x88FFFFFF), Color(0x33FFFFFF))) else SolidColor(Color(0x18FFFFFF)))
                        .border(1.dp, if (active) Color(0xAAFFFFFF) else Color(0x44FFFFFF), RoundedCornerShape(3.dp))
                        .clickable { onApp(a.id) },
                    contentAlignment = Alignment.Center
                ) { Icon(a.icon, a.label, tint = a.tint, modifier = Modifier.size(26.dp)) }
            }
        }
        Column(Modifier.padding(horizontal = 8.dp), horizontalAlignment = Alignment.CenterHorizontally) {
            Text(SimpleDateFormat("h:mm a", Locale.US).format(now), color = Color.White, fontSize = 12.sp)
            Text(SimpleDateFormat("dd/MM/yyyy", Locale.US).format(now), color = Color.White, fontSize = 11.sp)
        }
        // "Show desktop" strip like Win7
        Box(
            Modifier.width(12.dp).fillMaxHeight().background(Color(0x22FFFFFF)).border(1.dp, Color(0x33FFFFFF))
                .clickable(onClick = onDesktop)
        )
    }
}

data class StartEntry(val label: String, val icon: ImageVector, val tint: Color, val action: () -> Unit)

@Composable
fun StartMenu(left: List<StartEntry>, right: List<StartEntry>, onSearch: (String) -> Unit, onShutdown: () -> Unit, onDismiss: () -> Unit) {
    var q by remember { mutableStateOf("") }
    Box(Modifier.fillMaxSize().background(Color(0x22000000)).clickable(onClick = onDismiss)) {
        Row(
            Modifier.align(Alignment.BottomStart).padding(start = 2.dp).fillMaxWidth(0.94f).heightIn(max = 480.dp)
                .shadow(16.dp, RoundedCornerShape(topEnd = 8.dp, topStart = 8.dp))
                .clip(RoundedCornerShape(topEnd = 8.dp, topStart = 8.dp))
                .background(W7.startRight).border(1.dp, Color(0x88FFFFFF), RoundedCornerShape(topEnd = 8.dp, topStart = 8.dp))
                .clickable(enabled = false) {}
                .padding(6.dp)
        ) {
            Column(Modifier.weight(1.15f).fillMaxHeight().clip(RoundedCornerShape(4.dp)).background(Color.White)) {
                Column(Modifier.weight(1f).verticalScroll(rememberScrollState()).padding(4.dp)) {
                    left.forEach { e ->
                        Row(
                            Modifier.fillMaxWidth().clip(RoundedCornerShape(3.dp)).clickable { onDismiss(); e.action() }.padding(8.dp),
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Icon(e.icon, null, tint = e.tint, modifier = Modifier.size(28.dp))
                            Text(e.label, Modifier.padding(start = 10.dp), fontSize = 14.sp, color = W7.text)
                        }
                    }
                }
                HorizontalDivider(color = Color(0xFFD5DFEA))
                WinField(q, { q = it }, Modifier.fillMaxWidth().padding(6.dp), "Search programs and files", onDone = {
                    onDismiss(); onSearch(q)
                })
            }
            Column(Modifier.weight(1f).fillMaxHeight().padding(start = 6.dp)) {
                Box(
                    Modifier.padding(6.dp).size(56.dp).clip(RoundedCornerShape(6.dp)).background(W7.orb)
                        .border(2.dp, Color(0xCCFFFFFF), RoundedCornerShape(6.dp)),
                    contentAlignment = Alignment.Center
                ) { Text("FU", color = Color.White, fontWeight = FontWeight.Bold, fontSize = 20.sp) }
                Column(Modifier.weight(1f).verticalScroll(rememberScrollState())) {
                    right.forEach { e ->
                        Text(
                            e.label, Modifier.fillMaxWidth().clip(RoundedCornerShape(3.dp)).clickable { onDismiss(); e.action() }
                                .padding(horizontal = 8.dp, vertical = 9.dp),
                            color = Color.White, fontSize = 14.sp
                        )
                    }
                }
                Row(
                    Modifier.padding(6.dp).clip(RoundedCornerShape(3.dp)).background(W7.btnBlue)
                        .border(1.dp, Color(0xFF3C5A7A), RoundedCornerShape(3.dp)).clickable(onClick = onShutdown)
                        .padding(horizontal = 10.dp, vertical = 6.dp),
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Icon(Icons.Filled.PowerSettingsNew, null, tint = Color(0xFFC13A1F), modifier = Modifier.size(18.dp))
                    Text(" Shut down", fontSize = 13.sp, color = W7.text)
                }
            }
        }
    }
}

@Composable
fun DesktopIcon(label: String, icon: ImageVector, tint: Color, onClick: () -> Unit) {
    Column(
        Modifier.width(84.dp).clip(RoundedCornerShape(4.dp)).clickable(onClick = onClick).padding(6.dp),
        horizontalAlignment = Alignment.CenterHorizontally
    ) {
        Icon(icon, null, tint = tint, modifier = Modifier.size(44.dp))
        Text(
            label, color = Color.White, fontSize = 12.sp, maxLines = 2, textAlign = androidx.compose.ui.text.style.TextAlign.Center,
            style = TextStyle(shadow = androidx.compose.ui.graphics.Shadow(Color.Black, blurRadius = 4f))
        )
    }
}

// ---------------- File type icons ----------------

fun fileIcon(f: File): Pair<ImageVector, Color> = when (Fs.kind(f)) {
    Kind.FOLDER -> Icons.Filled.Folder to W7.folder
    Kind.VIDEO -> Icons.Filled.Movie to Color(0xFF6A3FB5)
    Kind.AUDIO -> Icons.Filled.MusicNote to Color(0xFFE0533D)
    Kind.IMAGE -> Icons.Filled.Image to Color(0xFF2E9E47)
    Kind.DOC -> if (f.extension.lowercase() == "pdf") Icons.Filled.PictureAsPdf to Color(0xFFD32F2F) else Icons.Filled.Description to Color(0xFF2B66C4)
    Kind.ARCHIVE -> Icons.Filled.Archive to Color(0xFFB8860B)
    Kind.APK -> Icons.Filled.Android to Color(0xFF3DDC84)
    Kind.OTHER -> Icons.Filled.InsertDriveFile to Color(0xFF8A8A8A)
}
