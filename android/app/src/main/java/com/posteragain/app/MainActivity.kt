package com.posteragain.app

import android.content.ContentValues
import android.content.Context
import android.content.Intent
import android.graphics.Bitmap
import android.graphics.BitmapFactory
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.provider.MediaStore
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.Image
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.lazy.LazyRow
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.verticalScroll
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.LinearProgressIndicator
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.rememberCoroutineScope
import androidx.compose.runtime.setValue
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.asImageBitmap
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.unit.dp
import com.chaquo.python.PyException
import com.chaquo.python.Python
import com.chaquo.python.android.AndroidPlatform
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.io.File

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        if (!Python.isStarted()) {
            Python.start(AndroidPlatform(this))
        }
        setContent {
            MaterialTheme { BatchScreen() }
        }
    }
}

private data class SessionInfo(
    val sid: String,
    val fields: List<String>,
    val headers: List<String>,
    val rowCount: Int,
    val matched: List<String>,
    val unmatched: List<String>,
)

/** 字段取值配置（手动绑定/固定值/前后缀；source: auto|column|fixed） */
private data class FieldConfig(
    val source: String,
    val value: String,
    val prefix: String = "",
    val suffix: String = "",
) {
    fun toJson(): org.json.JSONObject = org.json.JSONObject().apply {
        put("source", source); put("value", value)
        put("prefix", prefix); put("suffix", suffix)
    }

    fun describe(): String = when (source) {
        "column" -> "列『$value』" + deco()
        "fixed" -> "固定值『$value』" + deco()
        else -> "自动"
    }

    private fun deco(): String {
        val parts = listOfNotNull(
            prefix.takeIf { it.isNotEmpty() }?.let { "前缀『$it』" },
            suffix.takeIf { it.isNotEmpty() }?.let { "后缀『$it』" })
        return if (parts.isEmpty()) "" else "（${parts.joinToString("，")}）"
    }
}

@Composable
fun BatchScreen() {
    val context = LocalContext.current
    val scope = rememberCoroutineScope()

    var tplBytes by remember { mutableStateOf<ByteArray?>(null) }
    var dataBytes by remember { mutableStateOf<ByteArray?>(null) }
    var tplName by remember { mutableStateOf("") }
    var dataName by remember { mutableStateOf("") }
    var session by remember { mutableStateOf<SessionInfo?>(null) }
    var fieldConfigs by remember { mutableStateOf<Map<String, FieldConfig>>(emptyMap()) }
    var configDialogField by remember { mutableStateOf<String?>(null) }
    var restoredFromPrefs by remember { mutableStateOf(false) }
    var status by remember { mutableStateOf("① 选择模板  ② 选择数据  ③ 批量渲染") }
    var results by remember { mutableStateOf<List<String>>(emptyList()) }
    var progress by remember { mutableStateOf(-1 to 0) }   // done to total
    var busy by remember { mutableStateOf(false) }
    var selected by remember { mutableStateOf<String?>(null) }

    fun py() = Python.getInstance().getModule("poster_bridge")

    fun readUri(uri: Uri): ByteArray =
        context.contentResolver.openInputStream(uri)!!.use { it.readBytes() }

    // lastPathSegment 可能是截断的显示名/文档 ID，DISPLAY_NAME 才是真名
    fun queryName(uri: Uri): String =
        runCatching {
            context.contentResolver.query(uri, null, null, null, null)?.use { c ->
                val idx = c.getColumnIndex(android.provider.OpenableColumns.DISPLAY_NAME)
                if (idx >= 0 && c.moveToFirst()) c.getString(idx) else null
            }
        }.getOrNull() ?: (uri.lastPathSegment?.substringAfterLast('/') ?: "data")

    /** 把当前字段配置推送到 Python 会话（渲染时生效） */
    fun pushConfigs(sid: String, configs: Map<String, FieldConfig>) {
        scope.launch(Dispatchers.IO) {
            runCatching {
                val obj = org.json.JSONObject()
                session?.fields?.forEach { f ->
                    obj.put(f, (configs[f] ?: FieldConfig("auto", "")).toJson())
                }
                py().callAttr("apply_field_configs", sid, obj.toString())
            }.onFailure {
                withContext(Dispatchers.Main) { status = "应用字段配置失败: ${it.message}" }
            }
        }
    }

    fun openSession(t: ByteArray, d: ByteArray?, name: String) {
        scope.launch(Dispatchers.IO) {
            try {
                val old = session
                if (old != null) runCatching { py().callAttr("end_session", old.sid) }
                val r = org.json.JSONObject(
                    py().callAttr("start_session", t, d, name).toString())
                val info = SessionInfo(
                    sid = r.getString("session"),
                    fields = r.getJSONArray("fields").let { arr ->
                        List(arr.length()) { arr.getString(it) }
                    },
                    headers = r.optJSONArray("headers")?.let { arr ->
                        List(arr.length()) { arr.getString(it) }
                    } ?: emptyList(),
                    rowCount = r.getInt("row_count"),
                    matched = r.getJSONArray("matched").let { arr ->
                        List(arr.length()) { arr.getString(it) }
                    },
                    unmatched = r.getJSONArray("unmatched").let { arr ->
                        List(arr.length()) { arr.getString(it) }
                    },
                )
                // 套用记住的字段配置（模板名+字段指纹命中才生效）
                val restored = loadFieldConfigs(context, tplName, info.fields)
                if (restored.isNotEmpty()) {
                    fieldConfigs = restored
                    pushConfigs(info.sid, restored)
                } else {
                    fieldConfigs = emptyMap()
                }
                withContext(Dispatchers.Main) {
                    session = info
                    restoredFromPrefs = restored.isNotEmpty()
                    results = emptyList()
                    progress = -1 to 0
                    status = "模板就绪：${info.fields.size} 个占位符" +
                            (if (info.rowCount > 0) "；数据 ${info.rowCount} 行，匹配 ${info.matched.size} 个字段"
                             else "；未选数据") +
                            (if (restored.isNotEmpty()) "；已套用上次字段配置" else "")
                }
            } catch (e: PyException) {
                withContext(Dispatchers.Main) { status = "Python 错误: ${e.message}" }
            } catch (e: Exception) {
                withContext(Dispatchers.Main) { status = "错误: ${e.message}" }
            }
        }
    }

    val tplPicker = rememberLauncherForActivityResult(
        ActivityResultContracts.GetContent()
    ) { uri: Uri? ->
        if (uri != null) scope.launch(Dispatchers.IO) {
            try {
                val bytes = readUri(uri)
                val name = queryName(uri)
                withContext(Dispatchers.Main) {
                    tplBytes = bytes; tplName = name
                    status = "模板已读入（${bytes.size / 1024}KB）"
                }
                openSession(bytes, dataBytes, dataName)
            } catch (e: Exception) {
                withContext(Dispatchers.Main) { status = "读取模板失败: ${e.message}" }
            }
        }
    }

    val dataPicker = rememberLauncherForActivityResult(
        ActivityResultContracts.GetContent()
    ) { uri: Uri? ->
        if (uri != null) scope.launch(Dispatchers.IO) {
            try {
                val bytes = readUri(uri)
                val name = queryName(uri)
                withContext(Dispatchers.Main) { dataBytes = bytes; dataName = name }
                tplBytes?.let { openSession(it, bytes, name) }
                    ?: run { withContext(Dispatchers.Main) { status = "数据已读入，请先选模板" } }
            } catch (e: Exception) {
                withContext(Dispatchers.Main) { status = "读取数据失败: ${e.message}" }
            }
        }
    }

    fun runBatch() {
        val s = session ?: return
        if (s.rowCount <= 0 || busy) return
        busy = true
        // 渲染前确保当前字段配置已推送到会话并持久化
        pushConfigs(s.sid, fieldConfigs)
        saveFieldConfigs(context, tplName, s.fields, fieldConfigs)
        scope.launch(Dispatchers.IO) {
            try {
                val t0 = System.currentTimeMillis()
                val outRoot = File(context.filesDir, "batch").apply { mkdirs() }
                val outs = mutableListOf<String>()
                for (i in 0 until s.rowCount) {
                    val rowDir = File(outRoot, "row_$i")
                    val imgs = py().callAttr(
                        "render_index", s.sid, i, rowDir.absolutePath
                    ).asList().map { it.toString() }
                    outs += imgs
                    withContext(Dispatchers.Main) {
                        progress = (i + 1) to s.rowCount
                        selected = imgs.firstOrNull() ?: selected
                    }
                }
                val ms = (System.currentTimeMillis() - t0) / s.rowCount
                withContext(Dispatchers.Main) {
                    results = outs
                    status = "批量完成：${outs.size} 张，平均 ${ms}ms/张"
                }
            } catch (e: PyException) {
                withContext(Dispatchers.Main) { status = "Python 错误: ${e.message}" }
            } catch (e: Exception) {
                withContext(Dispatchers.Main) { status = "错误: ${e.message}" }
            } finally {
                withContext(Dispatchers.Main) { busy = false }
            }
        }
    }

    Column(
        Modifier.fillMaxSize().padding(16.dp).verticalScroll(rememberScrollState()),
        verticalArrangement = Arrangement.spacedBy(10.dp)
    ) {
        Row(horizontalArrangement = Arrangement.spacedBy(10.dp)) {
            Button(onClick = { tplPicker.launch("*/*") }, enabled = !busy) {
                Text(if (tplBytes == null) "① 选择模板" else "① 重选模板")
            }
            Button(onClick = { dataPicker.launch("*/*") }, enabled = !busy) {
                Text(if (dataBytes == null) "② 选择数据" else "② 重选数据")
            }
        }
        Row(horizontalArrangement = Arrangement.spacedBy(10.dp)) {
            Button(onClick = { runBatch() },
                enabled = !busy && session != null && session!!.rowCount > 0) {
                Text(if (busy) "渲染中…" else "③ 批量渲染（${session?.rowCount ?: 0} 行）")
            }
            Button(onClick = {
                session?.let { s ->
                    scope.launch(Dispatchers.IO) {
                        try {
                            val dir = File(context.filesDir, "preview")
                            val imgs = py().callAttr("render_preview", s.sid, dir.absolutePath)
                                .asList().map { it.toString() }
                            withContext(Dispatchers.Main) {
                                results = imgs; selected = imgs.firstOrNull()
                                status = "模板预览（未替换数据）"
                            }
                        } catch (e: Exception) {
                            withContext(Dispatchers.Main) { status = "错误: ${e.message}" }
                        }
                    }
                }
            }, enabled = !busy && session != null) { Text("预览模板") }
        }

        if (progress.first >= 0 && busy) {
            LinearProgressIndicator(
                progress = { progress.first.toFloat() / maxOf(1, progress.second) },
                modifier = Modifier.fillMaxWidth())
            Text("进度 ${progress.first}/${progress.second}", style = MaterialTheme.typography.bodySmall)
        }

        Text(status, style = MaterialTheme.typography.bodySmall)
        session?.let { s ->
            if (restoredFromPrefs) {
                Text("已套用上次保存的字段配置（可点击修改）",
                    style = MaterialTheme.typography.bodySmall,
                    color = MaterialTheme.colorScheme.primary)
            }
            Text("字段绑定（点击设置）：", style = MaterialTheme.typography.bodySmall)
            s.fields.forEach { f ->
                val auto = s.matched.firstOrNull { it.startsWith("$f →") }
                    ?.substringAfter(" → ")
                val cfg = fieldConfigs[f]
                val desc = cfg?.describe()
                    ?: (auto?.let { "列『$it』" } ?: "（未匹配，保留占位原样）")
                val missing = cfg == null && auto == null
                Text(
                    "• $f → $desc",
                    style = MaterialTheme.typography.bodySmall,
                    color = if (missing) MaterialTheme.colorScheme.error
                            else MaterialTheme.colorScheme.onSurface,
                    modifier = Modifier
                        .fillMaxWidth()
                        .clickable(enabled = !busy) { configDialogField = f }
                        .padding(vertical = 2.dp))
            }
        }

        if (results.isNotEmpty()) {
            Row(horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                Button(onClick = {
                    scope.launch(Dispatchers.IO) {
                        var ok = 0
                        results.forEach { if (savePngToGallery(context, it) != null) ok++ }
                        withContext(Dispatchers.Main) { status = "已保存 $ok/${results.size} 张到相册 Pictures/PosterAgain" }
                    }
                }) { Text("全部存相册") }
                Button(onClick = {
                    selected?.let { p ->
                        runCatching { sharePng(context, p) }
                            .onFailure { status = "分享失败: ${it.message}" }
                    }
                }) { Text("分享选中") }
            }
            LazyRow(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                items(results) { path ->
                    val bmp = remember(path) { BitmapFactory.decodeFile(path) }
                    bmp?.let {
                        Card(Modifier.width(140.dp).clickable { selected = path }) {
                            Image(it.asImageBitmap(), contentDescription = path,
                                modifier = Modifier.fillMaxWidth())
                        }
                    }
                }
            }
            selected?.let { p ->
                val bmp = remember(p) { BitmapFactory.decodeFile(p) }
                bmp?.let {
                    Text("选中预览：", style = MaterialTheme.typography.bodySmall)
                    Image(it.asImageBitmap(), contentDescription = "selected",
                        modifier = Modifier.fillMaxWidth())
                }
            }
        }

        configDialogField?.let { field ->
            val s = session ?: return@let
            FieldConfigDialog(
                field = field,
                headers = s.headers,
                autoMatched = s.matched.firstOrNull { it.startsWith("$field →") }
                    ?.substringAfter(" → "),
                initial = fieldConfigs[field],
                onDismiss = { configDialogField = null },
                onSave = { cfg ->
                    val newConfigs = if (cfg == null) fieldConfigs - field
                                    else fieldConfigs + (field to cfg)
                    fieldConfigs = newConfigs
                    configDialogField = null
                    pushConfigs(s.sid, newConfigs)
                    saveFieldConfigs(context, tplName, s.fields, newConfigs)
                })
        }
    }
}

private fun savePngToGallery(context: Context, path: String): Uri? {
    if (Build.VERSION.SDK_INT < Build.VERSION_CODES.Q) return null
    val values = ContentValues().apply {
        // 与渲染产物同名（如 1_陈春兰.png，与 PC 版导出命名规则一致）
        put(MediaStore.Images.Media.DISPLAY_NAME, File(path).name)
        put(MediaStore.Images.Media.MIME_TYPE, "image/png")
        put(MediaStore.Images.Media.RELATIVE_PATH,
            Environment.DIRECTORY_PICTURES + "/PosterAgain")
    }
    val uri = context.contentResolver.insert(
        MediaStore.Images.Media.EXTERNAL_CONTENT_URI, values) ?: return null
    context.contentResolver.openOutputStream(uri)?.use { out ->
        File(path).inputStream().use { it.copyTo(out) }
    }
    return uri
}

private fun sharePng(context: Context, path: String) {
    val uri = androidx.core.content.FileProvider.getUriForFile(
        context, "${context.packageName}.fileprovider", File(path))
    val intent = Intent(Intent.ACTION_SEND).apply {
        type = "image/png"
        putExtra(Intent.EXTRA_STREAM, uri)
        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    }
    context.startActivity(Intent.createChooser(intent, "分享海报"))
}

// ---------------- 字段配置持久化（记住配置）----------------
// 键 = 模板名 + 字段列表指纹：模板字段变了旧配置自动作废，不会错绑

private fun prefsKey(tplName: String, fields: List<String>): String =
    "fc_" + tplName + "_" + fields.sorted().joinToString("|").hashCode()

private fun loadFieldConfigs(context: Context, tplName: String,
                             fields: List<String>): Map<String, FieldConfig> {
    if (tplName.isEmpty() || fields.isEmpty()) return emptyMap()
    return runCatching {
        val json = context.getSharedPreferences("field_configs", Context.MODE_PRIVATE)
            .getString(prefsKey(tplName, fields), null) ?: return emptyMap()
        val obj = org.json.JSONObject(json)
        val out = mutableMapOf<String, FieldConfig>()
        for (k in obj.keys()) {
            val c = obj.getJSONObject(k)
            out[k] = FieldConfig(c.getString("source"), c.getString("value"),
                c.optString("prefix", ""), c.optString("suffix", ""))
        }
        out
    }.getOrDefault(emptyMap())
}

private fun saveFieldConfigs(context: Context, tplName: String,
                             fields: List<String>, configs: Map<String, FieldConfig>) {
    if (tplName.isEmpty() || fields.isEmpty()) return
    val obj = org.json.JSONObject()
    configs.forEach { (k, v) -> obj.put(k, v.toJson()) }
    context.getSharedPreferences("field_configs", Context.MODE_PRIVATE)
        .edit().putString(prefsKey(tplName, fields), obj.toString()).apply()
}

/** 字段设置对话框：数据源（自动/数据列/固定值）+ 前后缀 */
@Composable
private fun FieldConfigDialog(
    field: String,
    headers: List<String>,
    autoMatched: String?,
    initial: FieldConfig?,
    onDismiss: () -> Unit,
    onSave: (FieldConfig?) -> Unit,
) {
    var source by remember { mutableStateOf(initial?.source ?: "auto") }
    var column by remember {
        mutableStateOf(initial?.takeIf { it.source == "column" }?.value ?: "")
    }
    var fixed by remember {
        mutableStateOf(initial?.takeIf { it.source == "fixed" }?.value ?: "")
    }
    var prefix by remember { mutableStateOf(initial?.prefix ?: "") }
    var suffix by remember { mutableStateOf(initial?.suffix ?: "") }

    androidx.compose.material3.AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(field) },
        text = {
            Column(
                Modifier.verticalScroll(rememberScrollState()),
                verticalArrangement = Arrangement.spacedBy(8.dp)
            ) {
                Text(
                    if (autoMatched != null) "自动匹配 → $autoMatched" else "未自动匹配",
                    style = MaterialTheme.typography.bodySmall)
                Row(horizontalArrangement = Arrangement.spacedBy(6.dp)) {
                    listOf("自动" to "auto", "数据列" to "column", "固定值" to "fixed")
                        .forEach { (label, key) ->
                            androidx.compose.material3.Button(
                                onClick = { source = key },
                                colors = if (source == key)
                                    androidx.compose.material3.ButtonDefaults.buttonColors()
                                else androidx.compose.material3.ButtonDefaults.outlinedButtonColors(),
                                contentPadding = androidx.compose.foundation.layout.PaddingValues(
                                    horizontal = 10.dp, vertical = 4.dp)
                            ) { Text(label, style = MaterialTheme.typography.bodySmall) }
                        }
                }
                when (source) {
                    "column" -> {
                        if (headers.isEmpty()) {
                            Text("（数据未加载，无列可选）",
                                style = MaterialTheme.typography.bodySmall)
                        } else {
                            Text("选择数据列：", style = MaterialTheme.typography.bodySmall)
                            headers.forEach { h ->
                                Text(
                                    (if (h == column) "● " else "○ ") + h,
                                    style = MaterialTheme.typography.bodySmall,
                                    modifier = Modifier
                                        .fillMaxWidth()
                                        .clickable { column = h }
                                        .padding(vertical = 3.dp))
                            }
                        }
                    }
                    "fixed" -> {
                        androidx.compose.material3.OutlinedTextField(
                            value = fixed, onValueChange = { fixed = it },
                            label = { Text("固定值（应用到所有行）") },
                            singleLine = true)
                    }
                }
                androidx.compose.material3.OutlinedTextField(
                    value = prefix, onValueChange = { prefix = it },
                    label = { Text("前缀（可选）") }, singleLine = true)
                androidx.compose.material3.OutlinedTextField(
                    value = suffix, onValueChange = { suffix = it },
                    label = { Text("后缀（可选，如 元）") }, singleLine = true)
            }
        },
        confirmButton = {
            androidx.compose.material3.Button(onClick = {
                when {
                    source == "auto" -> onSave(null)
                    source == "column" && column.isNotEmpty() ->
                        onSave(FieldConfig("column", column, prefix, suffix))
                    source == "fixed" -> onSave(FieldConfig("fixed", fixed, prefix, suffix))
                    else -> onDismiss()
                }
            }) { Text("保存") }
        },
        dismissButton = {
            androidx.compose.material3.Button(onClick = onDismiss) { Text("取消") }
        }
    )
}
