package com.fu.fm

import android.content.Context
import android.net.Uri
import android.os.Handler
import android.os.Looper
import androidx.media3.common.MediaItem
import androidx.media3.common.util.UnstableApi
import androidx.media3.transformer.Composition
import androidx.media3.transformer.EditedMediaItem
import androidx.media3.transformer.ExportException
import androidx.media3.transformer.ExportResult
import androidx.media3.transformer.ProgressHolder
import androidx.media3.transformer.Transformer
import java.io.File

/** .m3u8 (HLS) -> .mp4. Main thread se call karo. */
@androidx.annotation.OptIn(UnstableApi::class)
object Convert {
    fun toMp4(
        ctx: Context, playlist: File, out: File,
        onProgress: (Int) -> Unit, onDone: (Boolean, String) -> Unit
    ) {
        val handler = Handler(Looper.getMainLooper())
        var finished = false
        lateinit var transformer: Transformer
        val holder = ProgressHolder()
        val poll = object : Runnable {
            override fun run() {
                if (finished) return
                if (transformer.getProgress(holder) == Transformer.PROGRESS_STATE_AVAILABLE) onProgress(holder.progress)
                handler.postDelayed(this, 500)
            }
        }
        transformer = Transformer.Builder(ctx)
            .addListener(object : Transformer.Listener {
                override fun onCompleted(composition: Composition, exportResult: ExportResult) {
                    finished = true
                    onDone(true, "")
                }

                override fun onError(composition: Composition, exportResult: ExportResult, exportException: ExportException) {
                    finished = true
                    out.delete()
                    onDone(false, exportException.message ?: "unknown error")
                }
            })
            .build()
        try {
            val item = EditedMediaItem.Builder(MediaItem.fromUri(Uri.fromFile(playlist))).build()
            transformer.start(item, out.absolutePath)
            handler.post(poll)
        } catch (e: Exception) {
            finished = true
            onDone(false, e.message ?: "start failed")
        }
    }
}
