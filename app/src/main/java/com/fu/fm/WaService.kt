package com.fu.fm

import android.app.NotificationChannel
import android.app.NotificationManager
import android.app.PendingIntent
import android.app.Service
import android.content.BroadcastReceiver
import android.content.Context
import android.content.Intent
import android.content.pm.ServiceInfo
import android.os.Build
import android.os.FileObserver
import android.os.Handler
import android.os.IBinder
import android.os.Looper
import androidx.core.app.NotificationCompat
import androidx.core.app.ServiceCompat
import androidx.core.content.ContextCompat
import java.io.File
import java.util.concurrent.Executors

/** WhatsApp Media folders ko live dekhta hai; har nayi file ki copy turant backup mai. */
class WaService : Service() {
    private val observers = HashMap<String, FileObserver>()
    private val exec = Executors.newSingleThreadExecutor()
    private val handler = Handler(Looper.getMainLooper())
    private val periodic = object : Runnable {
        override fun run() {
            exec.execute { safe { rewatch(); WaRecovery.sync() } }
            handler.postDelayed(this, 15 * 60_000L)
        }
    }

    override fun onBind(intent: Intent?): IBinder? = null

    override fun onCreate() {
        super.onCreate()
        val nm = getSystemService(NotificationManager::class.java)
        if (Build.VERSION.SDK_INT >= 26) {
            nm.createNotificationChannel(NotificationChannel(CH, "WhatsApp backup", NotificationManager.IMPORTANCE_MIN))
        }
        val pi = PendingIntent.getActivity(
            this, 0, Intent(this, MainActivity::class.java), PendingIntent.FLAG_IMMUTABLE
        )
        val n = NotificationCompat.Builder(this, CH)
            .setSmallIcon(android.R.drawable.ic_menu_save)
            .setContentTitle("WhatsApp recovery ON")
            .setContentText("Nayi photos/videos ki copy ban rahi hai")
            .setContentIntent(pi)
            .setOngoing(true)
            .setPriority(NotificationCompat.PRIORITY_MIN)
            .build()
        ServiceCompat.startForeground(
            this, 7, n,
            if (Build.VERSION.SDK_INT >= 29) ServiceInfo.FOREGROUND_SERVICE_TYPE_DATA_SYNC else 0
        )
    }

    override fun onStartCommand(intent: Intent?, flags: Int, startId: Int): Int {
        if (!WaRecovery.enabled(this)) { stopSelf(); return START_NOT_STICKY }
        handler.removeCallbacks(periodic)
        handler.post(periodic)
        return START_STICKY
    }

    override fun onDestroy() {
        handler.removeCallbacks(periodic)
        exec.execute { observers.values.forEach { it.stopWatching() }; observers.clear() }
        exec.shutdown()
        super.onDestroy()
    }

    private inline fun safe(block: () -> Unit) { try { block() } catch (_: Exception) {} }

    /** Exec thread pe hi chalao. Har wanted folder pe observer (FileObserver recursive nahi hota). */
    private fun rewatch() {
        observers.keys.filter { !File(it).isDirectory }.forEach { observers.remove(it)?.stopWatching() }
        fun walk(d: File) {
            watch(d)
            d.listFiles()?.forEach { if (it.isDirectory && WaRecovery.wantedDir(it)) walk(it) }
        }
        WaRecovery.sources().forEach { walk(it.second) }
    }

    @Suppress("DEPRECATION")
    private fun watch(dir: File) {
        val key = dir.absolutePath
        if (key in observers) return
        val mask = FileObserver.CLOSE_WRITE or FileObserver.MOVED_TO or FileObserver.CREATE
        val o = object : FileObserver(key, mask) {
            override fun onEvent(event: Int, path: String?) {
                if (path == null) return
                val f = File(dir, path)
                exec.execute {
                    safe {
                        if (f.isDirectory) {
                            if (WaRecovery.wantedDir(f)) {
                                watch(f)
                                f.walkTopDown().onEnter { it == f || WaRecovery.wantedDir(it) }
                                    .forEach { if (it.isDirectory) watch(it) else WaRecovery.save(it) }
                            }
                        } else if ((event and ALL_EVENTS) != CREATE) {
                            WaRecovery.save(f)
                        }
                    }
                }
            }
        }
        o.startWatching()
        observers[key] = o
    }

    companion object {
        private const val CH = "wa_backup"
        fun start(c: Context) = ContextCompat.startForegroundService(c, Intent(c, WaService::class.java))
        fun stop(c: Context) = c.stopService(Intent(c, WaService::class.java))
    }
}

/** Phone restart / app update ke baad service dobara chalao. */
class BootReceiver : BroadcastReceiver() {
    override fun onReceive(context: Context, intent: Intent) {
        if (WaRecovery.enabled(context)) try { WaService.start(context) } catch (_: Exception) {}
    }
}
