package com.posteragain.app

import android.content.Context
import android.net.Uri
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateListOf
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.setValue
import com.chaquo.python.PyException
import com.chaquo.python.Python
import kotlinx.coroutines.CancellationException
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.ExperimentalCoroutinesApi
import kotlinx.coroutines.Job
import kotlinx.coroutines.SupervisorJob
import kotlinx.coroutines.ensureActive
import kotlinx.coroutines.launch
import org.json.JSONArray
import org.json.JSONObject
import java.io.File

/** 路由：isTab 的三个是底部导航顶级页，切换会清空回退栈 */
enum class Route(val isTab: Boolean) {
    Home(true), Templates(true), Settings(true),
    Binding(false), Render(false), Preview(false)
}

data class SessionInfo(
    val sid: String,
    val fields: List<String>,
    val headers: List<String>,
    val rowCount: Int,
    val matched: List<String>,
    val unmatched: List<String>,
)

/** 字段取值配置（手动绑定/固定值/前后缀；source: auto|column|fixed） */
data class FieldConfig(
    val source: String,
    val value: String,
    val prefix: String = "",
    val suffix: String = "",
) {
    fun toJson(): JSONObject = JSONObject().apply {
        put("source", source); put("value", value)
        put("prefix", prefix); put("suffix", suffix)
    }

    fun chipText(): String = when (source) {
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

data class RecentTemplate(
    val hash: String,
    val name: String,
    val ts: Long,
    val count: Int,
)

/**
 * 全局状态容器：会话、渲染任务、最近模板、设置与导航。
 * 所有 Python 调用统一走单线程 gate，避免会话被并发读写。
 * Manifest 已锁定 orientation，无需 ViewModel 保转屏状态。
 */
@OptIn(ExperimentalCoroutinesApi::class)
class AppState(private val context: Context) {

    private val scope = CoroutineScope(SupervisorJob() + Dispatchers.Main.immediate)
    private val gate = Dispatchers.IO.limitedParallelism(1)

    private fun py() = Python.getInstance().getModule("poster_bridge")

    private val templatesDir get() = File(context.filesDir, "templates")
    private val thumbsDir get() = File(context.filesDir, "thumbs")

    // ---------------- 选择与会话 ----------------
    var tplBytes by mutableStateOf<ByteArray?>(null); private set
    var dataBytes by mutableStateOf<ByteArray?>(null); private set
    var tplName by mutableStateOf(""); private set
    var dataName by mutableStateOf(""); private set
    var tplHash by mutableStateOf<String?>(null); private set
    var session by mutableStateOf<SessionInfo?>(null); private set
    var fieldConfigs by mutableStateOf<Map<String, FieldConfig>>(emptyMap()); private set
    var restoredFromPrefs by mutableStateOf(false); private set
    var loadingSession by mutableStateOf(false); private set
    /** 一次性提示（snackbar），消费后置空 */
    var message by mutableStateOf<String?>(null)

    // ---------------- 渲染 ----------------
    enum class Phase { Idle, Rendering, Done }
    var phase by mutableStateOf(Phase.Idle); private set
    var progress by mutableStateOf(0 to 0)
    var results by mutableStateOf<List<String>>(emptyList()); private set
    var avgMs by mutableStateOf(0.0)
    var previewIndex by mutableStateOf(0)
    private var batchJob: Job? = null

    // ---------------- 最近模板与设置 ----------------
    var recents by mutableStateOf<List<RecentTemplate>>(emptyList()); private set
    var fontPolicy by mutableStateOf("auto"); private set      // auto | bundled
    var imageFormat by mutableStateOf("png"); private set      // png | jpg
    var exportSize by mutableStateOf("original"); private set  // original | 1080p | 2k

    // ---------------- 导航 ----------------
    var currentRoute by mutableStateOf(Route.Home); private set
    private val backStack = mutableStateListOf<Route>()
    val canGoBack: Boolean
        get() = backStack.isNotEmpty() ||
                (!currentRoute.isTab && currentRoute != Route.Home)

    init {
        loadSettings()
        loadRecents()
        applyFontPolicy()
    }

    fun navigate(route: Route) {
        if (route.isTab) {
            backStack.clear()
            currentRoute = route
        } else if (route != currentRoute) {
            backStack.add(currentRoute)
            currentRoute = route
        }
    }

    fun goBack() {
        if (backStack.isNotEmpty()) currentRoute = backStack.removeAt(backStack.lastIndex)
        else currentRoute = Route.Home
    }

    // ---------------- 文件选择 ----------------

    fun pickUri(uri: Uri, isTemplate: Boolean) {
        // Log.d 供 e2e_batch.py 验证选择回调已触发，勿删
        android.util.Log.d("E2E", "pickUri entry template=$isTemplate")
        scope.launch(gate) {
            try {
                val bytes = context.contentResolver.openInputStream(uri)!!.use { it.readBytes() }
                val name = queryDisplayName(context, uri)
                if (isTemplate) applyTemplate(bytes, name) else applyData(bytes, name)
            } catch (e: Exception) {
                android.util.Log.e("E2E", "pickUri failed", e)
                message = "读取文件失败: ${e.message}"
            }
        }
    }

    private fun applyData(bytes: ByteArray, name: String) {
        dataBytes = bytes
        dataName = name
        val t = tplBytes
        if (t != null && tplHash != null) openSessionInternal(t, bytes, tplName, tplHash!!)
        else message = "数据已读入（${bytes.size / 1024}KB），请先选择模板"
    }

    private fun applyTemplate(bytes: ByteArray, name: String) {
        val hash = sha16(bytes)
        // 落一份副本，供「最近模板」复用
        runCatching {
            templatesDir.mkdirs()
            File(templatesDir, "$hash.pptx").writeBytes(bytes)
        }
        touchRecent(hash, name)
        openSessionInternal(bytes, dataBytes, name, hash)
    }

    private fun openSessionInternal(t: ByteArray, d: ByteArray?, name: String, hash: String) {
        loadingSession = true
        try {
            session?.let { old -> runCatching { py().callAttr("end_session", old.sid) } }
            val r = JSONObject(py().callAttr("start_session", t, d, name).toString())
            android.util.Log.d("E2E", "start_session ok: $name")
            val info = SessionInfo(
                sid = r.getString("session"),
                fields = r.getJSONArray("fields").toStringList(),
                headers = r.optJSONArray("headers")?.toStringList() ?: emptyList(),
                rowCount = r.getInt("row_count"),
                matched = r.getJSONArray("matched").toStringList(),
                unmatched = r.getJSONArray("unmatched").toStringList(),
            )
            // 套用记住的字段配置（模板名+字段指纹命中才生效）
            val restored = parseConfigs(loadFieldConfigsJson(context, name, info.fields))
            if (restored.isNotEmpty()) pushConfigsDirect(info.sid, info.fields, restored)
            tplBytes = t; tplName = name; tplHash = hash
            session = info
            fieldConfigs = restored
            restoredFromPrefs = restored.isNotEmpty()
            phase = Phase.Idle
            results = emptyList()
            progress = 0 to 0
            ensureThumb(hash, info.sid)
        } catch (e: PyException) {
            android.util.Log.e("E2E", "openSession PyException", e)
            session = null
            message = "模板解析失败: ${e.message}"
        } catch (e: Exception) {
            android.util.Log.e("E2E", "openSession failed", e)
            message = "打开会话失败: ${e.message}"
        } finally {
            loadingSession = false
        }
    }

    /** 模板缩略图（render_preview 首图缩小存储），供首页/模板页卡片 */
    private fun ensureThumb(hash: String, sid: String) {
        val thumb = File(thumbsDir, "$hash.png")
        if (thumb.exists()) return
        runCatching {
            thumbsDir.mkdirs()
            val dir = File(context.cacheDir, "tpl_preview")
                .apply { deleteRecursively(); mkdirs() }
            val paths = py().callAttr("render_preview", sid, dir.absolutePath)
                .asList().map { it.toString() }
            val first = paths.firstOrNull() ?: return
            val bmp = decodeScaled(first, 320) ?: return
            thumb.outputStream().use { bmp.compress(android.graphics.Bitmap.CompressFormat.PNG, 100, it) }
        }
    }

    // ---------------- 字段配置 ----------------

    fun updateFieldConfig(field: String, cfg: FieldConfig?) {
        val s = session ?: return
        val newConfigs = if (cfg == null) fieldConfigs - field
                         else fieldConfigs + (field to cfg)
        fieldConfigs = newConfigs
        scope.launch(gate) {
            runCatching { pushConfigsDirect(s.sid, s.fields, newConfigs) }
            saveFieldConfigsJson(context, tplName, s.fields, configsToJson(newConfigs, s.fields))
        }
    }

    private fun pushConfigsDirect(sid: String, fields: List<String>, configs: Map<String, FieldConfig>) {
        py().callAttr("apply_field_configs", sid, configsToJson(configs, fields))
    }

    private fun parseConfigs(json: String?): Map<String, FieldConfig> {
        if (json == null) return emptyMap()
        return runCatching {
            val obj = JSONObject(json)
            val out = mutableMapOf<String, FieldConfig>()
            for (k in obj.keys()) {
                val c = obj.getJSONObject(k)
                out[k] = FieldConfig(c.getString("source"), c.getString("value"),
                    c.optString("prefix", ""), c.optString("suffix", ""))
            }
            out
        }.getOrDefault(emptyMap())
    }

    private fun configsToJson(configs: Map<String, FieldConfig>, fields: List<String>): String {
        val obj = JSONObject()
        fields.forEach { f -> obj.put(f, (configs[f] ?: FieldConfig("auto", "")).toJson()) }
        return obj.toString()
    }

    // ---------------- 渲染 ----------------

    fun startBatch() {
        val s = session ?: return
        if (phase == Phase.Rendering) return
        phase = Phase.Rendering
        results = emptyList()
        progress = 0 to s.rowCount
        avgMs = 0.0
        batchJob = scope.launch(gate) {
            val outRoot = File(context.filesDir, "batch").apply { deleteRecursively(); mkdirs() }
            val outs = mutableListOf<String>()
            val t0 = System.currentTimeMillis()
            var done = 0
            try {
                // 渲染前兜底推送一次当前字段配置
                runCatching { pushConfigsDirect(s.sid, s.fields, fieldConfigs) }
                for (i in 0 until s.rowCount) {
                    ensureActive()
                    val rowDir = File(outRoot, "row_$i")
                    val imgs = py().callAttr("render_index", s.sid, i, rowDir.absolutePath)
                        .asList().map { it.toString() }
                    outs += imgs
                    done = i + 1
                    progress = done to s.rowCount
                    avgMs = (System.currentTimeMillis() - t0).toDouble() / maxOf(1, done)
                }
                results = outs
                phase = Phase.Done
                message = "批量完成：${outs.size} 张，平均 ${avgMs.toInt()}ms/张"
            } catch (e: CancellationException) {
                results = outs
                phase = Phase.Done
                message = "已取消（完成 $done/${s.rowCount} 张）"
            } catch (e: PyException) {
                results = outs
                phase = Phase.Done
                message = "渲染中断: ${e.message}（已完成 $done 张）"
            } catch (e: Exception) {
                phase = Phase.Idle
                message = "渲染失败: ${e.message}"
            }
        }
    }

    fun cancelBatch() { batchJob?.cancel() }

    fun openPreview(index: Int) {
        previewIndex = index
        navigate(Route.Preview)
    }

    fun deleteResult(index: Int) {
        if (index !in results.indices) return
        runCatching { File(results[index]).delete() }
        results = results.toMutableList().apply { removeAt(index) }
        if (previewIndex > results.lastIndex) previewIndex = maxOf(0, results.lastIndex)
    }

    // ---------------- 保存 / 分享（格式与尺寸跟随设置） ----------------

    private fun scaleMaxDim(): Int? = when (exportSize) {
        "1080p" -> 1080
        "2k" -> 2048
        else -> null
    }

    fun saveOneToGallery(path: String) {
        scope.launch(gate) {
            val uri = runCatching { saveImageToGallery(context, path, imageFormat, scaleMaxDim()) }
                .getOrNull()
            message = if (uri != null) "已保存到相册 Pictures/PosterAgain" else "保存失败"
        }
    }

    fun saveAllToGallery() {
        val paths = results.toList()
        if (paths.isEmpty()) return
        scope.launch(gate) {
            var ok = 0
            paths.forEach { p ->
                if (runCatching { saveImageToGallery(context, p, imageFormat, scaleMaxDim()) }
                        .getOrNull() != null) ok++
            }
            message = "已保存 $ok/${paths.size} 张到相册 Pictures/PosterAgain"
        }
    }

    // ---------------- 最近模板 ----------------

    fun thumbFile(hash: String): File = File(thumbsDir, "$hash.png")

    fun loadRecentTemplate(r: RecentTemplate) {
        scope.launch(gate) {
            val f = File(templatesDir, "${r.hash}.pptx")
            if (!f.exists()) {
                removeRecent(r)
                message = "模板文件已失效，已从列表移除"
                return@launch
            }
            applyTemplate(f.readBytes(), r.name)
            navigate(Route.Home)
        }
    }

    fun deleteRecent(r: RecentTemplate) {
        removeRecent(r)
        runCatching {
            File(templatesDir, "${r.hash}.pptx").delete()
            File(thumbsDir, "${r.hash}.png").delete()
        }
    }

    private fun touchRecent(hash: String, name: String) {
        val list = recents.toMutableList()
        val existing = list.indexOfFirst { it.hash == hash }
        if (existing >= 0) {
            val e = list.removeAt(existing)
            list.add(0, e.copy(ts = System.currentTimeMillis(), count = e.count + 1))
        } else {
            list.add(0, RecentTemplate(hash, name, System.currentTimeMillis(), 1))
        }
        // 超出上限的旧条目连同文件一起清理
        while (list.size > MAX_RECENTS) {
            val dropped = list.removeAt(list.lastIndex)
            runCatching {
                File(templatesDir, "${dropped.hash}.pptx").delete()
                File(thumbsDir, "${dropped.hash}.png").delete()
            }
        }
        recents = list
        persistRecents()
    }

    private fun removeRecent(r: RecentTemplate) {
        recents = recents.filterNot { it.hash == r.hash }
        persistRecents()
    }

    private fun persistRecents() {
        val arr = JSONArray()
        recents.forEach { r ->
            arr.put(JSONObject().apply {
                put("hash", r.hash); put("name", r.name)
                put("ts", r.ts); put("count", r.count)
            })
        }
        context.getSharedPreferences("recents", Context.MODE_PRIVATE)
            .edit().putString("list", arr.toString()).apply()
    }

    private fun loadRecents(): List<RecentTemplate> {
        val json = context.getSharedPreferences("recents", Context.MODE_PRIVATE)
            .getString("list", null) ?: return emptyList()
        return runCatching {
            val arr = JSONArray(json)
            (0 until arr.length()).map { i ->
                val o = arr.getJSONObject(i)
                RecentTemplate(o.getString("hash"), o.getString("name"),
                    o.getLong("ts"), o.getInt("count"))
            }
        }.getOrDefault(emptyList())
    }

    // ---------------- 设置 ----------------

    fun updateFontPolicy(v: String) {
        fontPolicy = v
        persistSettings()
        applyFontPolicy()
    }

    fun updateImageFormat(v: String) { imageFormat = v; persistSettings() }

    fun updateExportSize(v: String) { exportSize = v; persistSettings() }

    private fun persistSettings() {
        context.getSharedPreferences("settings", Context.MODE_PRIVATE).edit()
            .putString("font_policy", fontPolicy)
            .putString("image_format", imageFormat)
            .putString("export_size", exportSize)
            .apply()
    }

    private fun loadSettings() {
        val p = context.getSharedPreferences("settings", Context.MODE_PRIVATE)
        fontPolicy = p.getString("font_policy", "auto") ?: "auto"
        imageFormat = p.getString("image_format", "png") ?: "png"
        exportSize = p.getString("export_size", "original") ?: "original"
    }

    /** 字体策略经环境变量注入 render_core（font_registry 惰性读取，随时生效） */
    private fun applyFontPolicy() {
        val value = if (fontPolicy == "bundled") "bundled" else "auto"
        scope.launch(gate) {
            runCatching { py().callAttr("set_env", "POSTERAGAIN_FONT_POLICY", value) }
        }
    }

    companion object { private const val MAX_RECENTS = 12 }
}

private fun org.json.JSONArray.toStringList(): List<String> =
    (0 until length()).map { getString(it) }
