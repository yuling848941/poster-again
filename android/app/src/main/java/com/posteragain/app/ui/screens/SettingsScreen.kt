package com.posteragain.app.ui.screens

import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.size
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.outlined.ArrowBack
import androidx.compose.material.icons.outlined.ChevronRight
import androidx.compose.material.icons.outlined.Description
import androidx.compose.material.icons.outlined.Folder
import androidx.compose.material.icons.outlined.FontDownload
import androidx.compose.material.icons.outlined.Image
import androidx.compose.material.icons.outlined.Info
import androidx.compose.material.icons.outlined.PhotoSizeSelectLarge
import androidx.compose.material3.AlertDialog
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.RadioButton
import androidx.compose.material3.Surface
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.vector.ImageVector
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import com.posteragain.app.AppState
import com.posteragain.app.BuildConfig
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

/** 设置：生成（字体方案/图片格式/导出尺寸）、存储、关于 */
@Composable
fun SettingsScreen(state: AppState) {
    var dialog by remember { mutableStateOf<String?>(null) }

    Column(Modifier.fillMaxSize()) {
        // 顶级 tab 页，顶部仅标题、无返回箭头
        Box(Modifier.fillMaxWidth().height(56.dp), contentAlignment = Alignment.Center) {
            Text("设置", style = MaterialTheme.typography.titleLarge)
        }

        Column(
            Modifier.fillMaxSize().verticalScroll(rememberScrollState()).padding(16.dp),
            verticalArrangement = Arrangement.spacedBy(16.dp)
        ) {
            SettingsGroup("生成") {
                SettingRow(Icons.Outlined.FontDownload, "字体方案",
                    if (state.fontPolicy == "bundled") "内置思源黑体" else "跟随系统",
                    onClick = { dialog = "font" })
                SettingRow(Icons.Outlined.Image, "图片格式",
                    state.imageFormat.uppercase(),
                    onClick = { dialog = "format" })
                SettingRow(Icons.Outlined.PhotoSizeSelectLarge, "导出尺寸",
                    when (state.exportSize) {
                        "1080p" -> "1080P"
                        "2k" -> "2K"
                        else -> "原始"
                    },
                    onClick = { dialog = "size" })
            }
            SettingsGroup("存储") {
                SettingRow(Icons.Outlined.Folder, "保存位置",
                    "Pictures/PosterAgain", onClick = null)
            }
            SettingsGroup("关于") {
                SettingRow(Icons.Outlined.Info, "版本", appVersion(), onClick = null)
                SettingRow(Icons.Outlined.Description, "开源字体声明",
                    "SIL OFL 1.1", onClick = { dialog = "ofl" })
            }
        }
    }

    when (dialog) {
        "font" -> SingleChoiceDialog(
            title = "字体方案",
            options = listOf("跟随系统" to "auto", "内置思源黑体" to "bundled"),
            selected = state.fontPolicy,
            onSelect = { state.updateFontPolicy(it) },
            onDismiss = { dialog = null })
        "format" -> SingleChoiceDialog(
            title = "图片格式",
            options = listOf("PNG" to "png", "JPG（体积更小）" to "jpg"),
            selected = state.imageFormat,
            onSelect = { state.updateImageFormat(it) },
            onDismiss = { dialog = null })
        "size" -> SingleChoiceDialog(
            title = "导出尺寸",
            options = listOf("原始" to "original", "1080P" to "1080p", "2K" to "2k"),
            selected = state.exportSize,
            onSelect = { state.updateExportSize(it) },
            onDismiss = { dialog = null })
        "ofl" -> AlertDialog(
            onDismissRequest = { dialog = null },
            title = { Text("开源字体声明") },
            text = {
                Text("应用内置的思源黑体（Source Han Sans）与思源宋体（Source Han Serif）" +
                        "子集由 Adobe 发布，遵循 SIL Open Font License 1.1 授权，" +
                        "可自由使用、研究与修改。许可证全文见：" +
                        "https://openfontlicense.org")
            },
            confirmButton = {
                TextButton(onClick = { dialog = null }) { Text("知道了") }
            })
    }
}

@Composable
private fun SettingsGroup(title: String, content: @Composable () -> Unit) {
    Column {
        Text(title, style = MaterialTheme.typography.labelLarge,
            color = MaterialTheme.colorScheme.primary,
            modifier = Modifier.padding(start = 8.dp, bottom = 6.dp))
        Surface(
            shape = RoundedCornerShape(16.dp),
            color = MaterialTheme.colorScheme.surfaceContainerLowest,
            border = androidx.compose.foundation.BorderStroke(
                1.dp, MaterialTheme.colorScheme.outlineVariant.copy(alpha = 0.4f))
        ) {
            Column { content() }
        }
    }
}

@Composable
private fun SettingRow(
    icon: ImageVector, title: String, value: String, onClick: (() -> Unit)?
) {
    Row(
        Modifier.fillMaxWidth()
            .then(if (onClick != null)
                Modifier.clickable(onClick = onClick) else Modifier)
            .padding(horizontal = 16.dp, vertical = 14.dp),
        verticalAlignment = Alignment.CenterVertically
    ) {
        Icon(icon, contentDescription = null,
            tint = MaterialTheme.colorScheme.onSurfaceVariant,
            modifier = Modifier.size(22.dp))
        Spacer(Modifier.width(14.dp))
        Text(title, style = MaterialTheme.typography.bodyLarge,
            fontWeight = FontWeight.Medium, modifier = Modifier.weight(1f))
        Text(value, style = MaterialTheme.typography.bodyMedium,
            color = MaterialTheme.colorScheme.onSurfaceVariant)
        if (onClick != null) {
            Spacer(Modifier.width(2.dp))
            Icon(Icons.Outlined.ChevronRight, contentDescription = null,
                tint = MaterialTheme.colorScheme.outline, modifier = Modifier.size(18.dp))
        }
    }
}

@Composable
private fun SingleChoiceDialog(
    title: String,
    options: List<Pair<String, String>>,
    selected: String,
    onSelect: (String) -> Unit,
    onDismiss: () -> Unit
) {
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(title) },
        text = {
            Column {
                options.forEach { (label, key) ->
                    Row(
                        verticalAlignment = Alignment.CenterVertically,
                        modifier = Modifier.fillMaxWidth()
                    ) {
                        RadioButton(
                            selected = key == selected,
                            onClick = { onSelect(key); onDismiss() })
                        Text(label, style = MaterialTheme.typography.bodyLarge)
                    }
                }
            }
        },
        confirmButton = {
            TextButton(onClick = onDismiss) { Text("取消") }
        }
    )
}

@Composable
private fun appVersion(): String {
    val context = LocalContext.current
    return remember {
        runCatching {
            val pi = context.packageManager.getPackageInfo(context.packageName, 0)
            "${pi.versionName}（构建于 ${
                SimpleDateFormat("yyyy-MM-dd", Locale.getDefault())
                    .format(Date(pi.lastUpdateTime))
            }）"
        }.getOrDefault("v${BuildConfig.VERSION_NAME}")
    }
}
