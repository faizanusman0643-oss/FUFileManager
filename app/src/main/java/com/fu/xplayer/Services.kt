package com.fu.xplayer

import android.app.Notification
import android.app.NotificationChannel
import android.app.NotificationManager
import android.app.PendingIntent
import android.app.Service
import android.content.Context
import android.content.Intent
import android.content.pm.ServiceInfo
import android.os.Build
import android.os.IBinder
import androidx.annotation.OptIn
import androidx.core.app.NotificationCompat
import androidx.core.app.ServiceCompat
import androidx.media3.common.AudioAttributes
import androidx.media3.common.C
import androidx.media3.common.util.UnstableApi
import androidx.media3.exoplayer.ExoPlayer
import androidx.media3.session.MediaSession
import androidx.media3.session.MediaSessionService
import kotlinx.coroutines.*

/** Keeps downloads alive in background with a progress notification. */
class DlService : Service() {
    private val scope = CoroutineScope(SupervisorJob() + Dispatchers.Main)

    override fun onBind(intent: Intent?): IBinder? = null

    override fun onStartCommand(intent: Intent?, flags: Int, startId: Int): Int {
        channel(this)
        val n = build("Downloads", "Shuru ho raha hai...", 0, true)
        try {
            if (Build.VERSION.SDK_INT >= 29) ServiceCompat.startForeground(this, 7, n, ServiceInfo.FOREGROUND_SERVICE_TYPE_DATA_SYNC)
            else startForeground(7, n)
        } catch (_: Exception) {}
        scope.coroutineContext.cancelChildren()
        scope.launch {
            var idle = 0
            while (isActive) {
                delay(1000)
                val act = Dl.tasks.filter { it.state == DState.RUNNING || it.state == DState.CONVERTING }
                if (act.isEmpty()) { if (++idle >= 3) break; continue }
                idle = 0
                val sp = act.sumOf { it.speed }
                val fr = act.mapNotNull { it.fraction }.let { if (it.isEmpty()) -1f else it.average().toFloat() }
                val title = if (act.size == 1) act[0].name else "${act.size} downloads chal rahe hain"
                val nm = getSystemService(NotificationManager::class.java)
                try { nm.notify(7, build(title, "${fmtSize(sp)}/s", (fr * 100).toInt(), fr < 0)) } catch (_: Exception) {}
            }
            ServiceCompat.stopForeground(this@DlService, ServiceCompat.STOP_FOREGROUND_REMOVE)
            stopSelf()
        }
        return START_NOT_STICKY
    }

    private fun build(title: String, text: String, p: Int, indet: Boolean): Notification {
        val pi = PendingIntent.getActivity(
            this, 0, Intent(this, MainActivity::class.java).putExtra("screen", "downloads"),
            PendingIntent.FLAG_IMMUTABLE or PendingIntent.FLAG_UPDATE_CURRENT
        )
        return NotificationCompat.Builder(this, "dl")
            .setSmallIcon(android.R.drawable.stat_sys_download)
            .setContentTitle(title).setContentText(text)
            .setProgress(100, p.coerceIn(0, 100), indet)
            .setOngoing(true).setOnlyAlertOnce(true).setContentIntent(pi)
            .build()
    }

    override fun onDestroy() { scope.cancel(); super.onDestroy() }

    companion object {
        fun channel(ctx: Context) {
            if (Build.VERSION.SDK_INT >= 26) {
                ctx.getSystemService(NotificationManager::class.java).createNotificationChannel(
                    NotificationChannel("dl", "Downloads", NotificationManager.IMPORTANCE_LOW)
                )
            }
        }
    }
}

/** Shared player used by the UI and by the background media notification. */
@OptIn(UnstableApi::class)
object Players {
    private var p: ExoPlayer? = null
    fun get(ctx: Context): ExoPlayer = p ?: ExoPlayer.Builder(ctx.applicationContext)
        .setAudioAttributes(AudioAttributes.Builder().setUsage(C.USAGE_MEDIA).setContentType(C.AUDIO_CONTENT_TYPE_MOVIE).build(), true)
        .setHandleAudioBecomingNoisy(true)
        .setSeekBackIncrementMs(10_000).setSeekForwardIncrementMs(10_000)
        .build().also { p = it }
}

class PlaybackService : MediaSessionService() {
    private var session: MediaSession? = null

    override fun onCreate() {
        super.onCreate()
        val pi = PendingIntent.getActivity(
            this, 1, Intent(this, MainActivity::class.java).putExtra("screen", "player"),
            PendingIntent.FLAG_IMMUTABLE or PendingIntent.FLAG_UPDATE_CURRENT
        )
        session = MediaSession.Builder(this, Players.get(this)).setSessionActivity(pi).build()
    }

    override fun onGetSession(controllerInfo: MediaSession.ControllerInfo): MediaSession? = session

    override fun onTaskRemoved(rootIntent: Intent?) {
        val pl = session?.player
        if (pl == null || !pl.playWhenReady || pl.mediaItemCount == 0) stopSelf()
    }

    override fun onDestroy() {
        session?.release(); session = null
        super.onDestroy()
    }
}

/** Simple "task complete" notifications (downloads, extract, copy). */
object Notify {
    fun show(ctx: Context, id: Int, title: String, text: String, file: java.io.File?) {
        try {
            if (Build.VERSION.SDK_INT >= 26) {
                ctx.getSystemService(NotificationManager::class.java).createNotificationChannel(
                    NotificationChannel("done", "Complete alerts", NotificationManager.IMPORTANCE_HIGH)
                )
            }
            val open = Intent(ctx, MainActivity::class.java).addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
            if (file != null) open.putExtra("open", file.absolutePath) else open.putExtra("screen", "downloads")
            val pi = PendingIntent.getActivity(ctx, id, open, PendingIntent.FLAG_IMMUTABLE or PendingIntent.FLAG_UPDATE_CURRENT)
            val n = NotificationCompat.Builder(ctx, "done")
                .setSmallIcon(android.R.drawable.stat_sys_download_done)
                .setContentTitle(title).setContentText(text)
                .setStyle(NotificationCompat.BigTextStyle().bigText(text))
                .setPriority(NotificationCompat.PRIORITY_HIGH)
                .setAutoCancel(true).setContentIntent(pi).build()
            ctx.getSystemService(NotificationManager::class.java).notify(id, n)
        } catch (_: Exception) {}
    }
}
