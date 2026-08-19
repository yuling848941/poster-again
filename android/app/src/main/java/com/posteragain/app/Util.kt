package com.posteragain.app

import android.content.ContentValues
import android.content.Context
import android.content.Intent
import android.graphics.Bitmap
import android.graphics.BitmapFactory
import android.net.Uri
import android.os.Build
import android.os.Environment
import android.provider.MediaStore
import android.provider.OpenableColumns
import androidx.core.content.FileProvider
import java.io.File
import java.security.MessageDigest

// ---------------- 文件与图片工具 ----------------

/** lastPathSegment 可能是截断的显示名/文档 ID，DISPLAY_NAME 才是真名 */
fun queryDisplayName(context: Context, uri: Uri): String =
    runCatching {
        context.contentResolver.query(uri, null, null, null, null)?.use { c ->
            val idx = c.getColumnIndex(OpenableColumns.DISPLAY_NAME)
            if (idx >= 0 && c.moveToFirst()) c.getString(idx) else null
        }
    }.getOrNull() ?: (uri.lastPathSegment?.substringAfterLast('/') ?: "data")

fun sha16(bytes: ByteArray): String =
    MessageDigest.getInstance("SHA-256").digest(bytes)
        .joinToString("") { "%02x".format(it) }.take(16)

/** 按需解码为缩略图（最长边不超过 maxDim），供网格/最近使用卡片使用 */
fun decodeScaled(path: String, maxDim: Int): Bitmap? {
    val opts = BitmapFactory.Options().apply { inJustDecodeBounds = true }
    BitmapFactory.decodeFile(path, opts)
    if (opts.outWidth <= 0 || opts.outHeight <= 0) return null
    var sample = 1
    while (maxOf(opts.outWidth, opts.outHeight) / (sample * 2) >= maxDim) sample *= 2
    return BitmapFactory.decodeFile(path, BitmapFactory.Options().apply {
        inSampleSize = sample
    })
}

/** 缩放到指定最长边（scaleMaxDim 为 null 时保持原始尺寸） */
private fun Bitmap.scaleTo(scaleMaxDim: Int?): Bitmap {
    if (scaleMaxDim == null) return this
    val longSide = maxOf(width, height)
    if (longSide <= scaleMaxDim) return this
    val k = scaleMaxDim.toFloat() / longSide
    return Bitmap.createScaledBitmap(this, (width * k).toInt(), (height * k).toInt(), true)
}

/**
 * 保存到系统相册 Pictures/PosterAgain。
 * format: "png" | "jpg"；scaleMaxDim: null(原始) | 1080 | 2048。
 * png+原始尺寸走文件直拷快速路径，其余解码后按设置重编码。
 */
fun saveImageToGallery(
    context: Context, path: String, format: String, scaleMaxDim: Int?
): Uri? {
    if (Build.VERSION.SDK_INT < Build.VERSION_CODES.Q) return null
    val isFastPath = format == "png" && scaleMaxDim == null
    val ext = if (format == "jpg") "jpg" else "png"
    val mime = if (format == "jpg") "image/jpeg" else "image/png"
    val base = File(path).nameWithoutExtension
    val values = ContentValues().apply {
        // 与渲染产物同名（如 1_陈春兰.png，与 PC 版导出命名规则一致）
        put(MediaStore.Images.Media.DISPLAY_NAME, "$base.$ext")
        put(MediaStore.Images.Media.MIME_TYPE, mime)
        put(MediaStore.Images.Media.RELATIVE_PATH,
            Environment.DIRECTORY_PICTURES + "/PosterAgain")
    }
    val uri = context.contentResolver.insert(
        MediaStore.Images.Media.EXTERNAL_CONTENT_URI, values) ?: return null
    context.contentResolver.openOutputStream(uri)?.use { out ->
        if (isFastPath) {
            File(path).inputStream().use { it.copyTo(out) }
        } else {
            val bmp = BitmapFactory.decodeFile(path)?.scaleTo(scaleMaxDim) ?: run {
                context.contentResolver.delete(uri, null, null)
                return null
            }
            if (format == "jpg") bmp.compress(Bitmap.CompressFormat.JPEG, 92, out)
            else bmp.compress(Bitmap.CompressFormat.PNG, 100, out)
        }
    } ?: return null
    return uri
}

fun fileProviderUri(context: Context, path: String): Uri =
    FileProvider.getUriForFile(
        context, "${context.packageName}.fileprovider", File(path))

fun shareImage(context: Context, path: String) {
    val intent = Intent(Intent.ACTION_SEND).apply {
        type = "image/png"
        putExtra(Intent.EXTRA_STREAM, fileProviderUri(context, path))
        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    }
    context.startActivity(Intent.createChooser(intent, "分享海报"))
}

fun shareImages(context: Context, paths: List<String>) {
    val uris = ArrayList(paths.map { fileProviderUri(context, it) })
    val intent = Intent(Intent.ACTION_SEND_MULTIPLE).apply {
        type = "image/png"
        putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris)
        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    }
    context.startActivity(Intent.createChooser(intent, "分享 ${paths.size} 张海报"))
}

/** 设为壁纸（系统裁剪/选择器，ACTION_ATTACH_DATA） */
fun setAsWallpaper(context: Context, path: String) {
    val intent = Intent(Intent.ACTION_ATTACH_DATA).apply {
        setDataAndType(fileProviderUri(context, path), "image/png")
        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    }
    context.startActivity(Intent.createChooser(intent, "设为壁纸"))
}

// ---------------- 字段配置持久化（记住配置）----------------
// 键 = 模板名 + 字段列表指纹：模板字段变了旧配置自动作废，不会错绑

fun fieldConfigsPrefsKey(tplName: String, fields: List<String>): String =
    "fc_" + tplName + "_" + fields.sorted().joinToString("|").hashCode()

fun loadFieldConfigsJson(context: Context, tplName: String, fields: List<String>): String? {
    if (tplName.isEmpty() || fields.isEmpty()) return null
    return context.getSharedPreferences("field_configs", Context.MODE_PRIVATE)
        .getString(fieldConfigsPrefsKey(tplName, fields), null)
}

fun saveFieldConfigsJson(context: Context, tplName: String, fields: List<String>, json: String) {
    if (tplName.isEmpty() || fields.isEmpty()) return
    context.getSharedPreferences("field_configs", Context.MODE_PRIVATE)
        .edit().putString(fieldConfigsPrefsKey(tplName, fields), json).apply()
}
