package com.posteragain.app.ui.screens

import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.heightIn
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.size
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.outlined.ArrowBack
import androidx.compose.material.icons.outlined.ChevronRight
import androidx.compose.material.icons.outlined.Info
import androidx.compose.material3.Button
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.ModalBottomSheet
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.RadioButton
import androidx.compose.material3.SegmentedButton
import androidx.compose.material3.SegmentedButtonDefaults
import androidx.compose.material3.SingleChoiceSegmentedButtonRow
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
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import com.posteragain.app.AppState
import com.posteragain.app.FieldConfig
import com.posteragain.app.Route
import com.posteragain.app.ui.theme.PosterTheme

/** 字段绑定：占位符列表 + 点击弹出的字段配置底部抽屉 */
@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun BindingScreen(state: AppState) {
    val s = state.session ?: run {
        EmptyState("尚未选择模板，请回到首页开始"); return
    }
    var sheetField by remember { mutableStateOf<String?>(null) }

    Column(Modifier.fillMaxSize()) {
        // 顶栏
        Row(
            Modifier.fillMaxWidth().height(56.dp).padding(horizontal = 4.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            IconButton(onClick = { state.goBack() }) {
                Icon(Icons.AutoMirrored.Outlined.ArrowBack, contentDescription = "返回")
            }
            Text("字段绑定",
                style = MaterialTheme.typography.titleLarge,
                modifier = Modifier.weight(1f))
            TextButton(
                onClick = { state.navigate(Route.Render) },
                enabled = s.rowCount > 0 && state.phase != AppState.Phase.Rendering
            ) { Text("开始渲染", fontWeight = FontWeight.SemiBold) }
        }

        Column(
            Modifier.weight(1f).verticalScroll(rememberScrollState()).padding(horizontal = 16.dp)
        ) {
            // 概要信息行
            Text(
                "模板 ${s.fields.size} 个占位符 · 数据 ${s.rowCount} 行 · 自动匹配 ${s.matched.size} 个",
                style = MaterialTheme.typography.bodySmall,
                color = MaterialTheme.colorScheme.onSurfaceVariant)
            Spacer(Modifier.height(10.dp))

            // 已套用记忆配置的提示条
            if (state.restoredFromPrefs) {
                Surface(
                    shape = RoundedCornerShape(12.dp),
                    color = MaterialTheme.colorScheme.secondaryContainer,
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Row(Modifier.padding(12.dp), verticalAlignment = Alignment.CenterVertically) {
                        Icon(Icons.Outlined.Info, contentDescription = null,
                            modifier = Modifier.size(16.dp),
                            tint = MaterialTheme.colorScheme.onSecondaryContainer)
                        Spacer(Modifier.width(8.dp))
                        Text("已套用上次保存的字段配置，点击任意字段可修改",
                            style = MaterialTheme.typography.bodySmall,
                            color = MaterialTheme.colorScheme.onSecondaryContainer)
                    }
                }
                Spacer(Modifier.height(10.dp))
            }

            // 字段卡片列表
            Column(verticalArrangement = Arrangement.spacedBy(8.dp)) {
                s.fields.forEach { field ->
                    FieldCard(
                        field = field,
                        config = state.fieldConfigs[field],
                        autoMatched = s.matched.firstOrNull { it.startsWith("$field →") }
                            ?.substringAfter(" → "),
                        onClick = { sheetField = field })
                }
            }
            Spacer(Modifier.height(8.dp))
            Text("未绑定的字段将保留模板占位原样",
                style = MaterialTheme.typography.bodySmall,
                color = MaterialTheme.colorScheme.onSurfaceVariant)
            Spacer(Modifier.height(24.dp))
        }
    }

    // 字段配置抽屉
    sheetField?.let { field ->
        ModalBottomSheet(onDismissRequest = { sheetField = null }) {
            FieldConfigSheet(
                field = field,
                headers = s.headers,
                autoMatched = s.matched.firstOrNull { it.startsWith("$field →") }
                    ?.substringAfter(" → "),
                initial = state.fieldConfigs[field],
                onSave = { cfg ->
                    state.updateFieldConfig(field, cfg)
                    sheetField = null
                },
                onDismiss = { sheetField = null })
        }
    }
}

@Composable
private fun FieldCard(
    field: String,
    config: FieldConfig?,
    autoMatched: String?,
    onClick: () -> Unit
) {
    val unmatched = config == null && autoMatched == null
    Surface(
        shape = RoundedCornerShape(14.dp),
        color = MaterialTheme.colorScheme.surfaceContainerLowest,
        border = androidx.compose.foundation.BorderStroke(
            1.dp,
            if (unmatched) MaterialTheme.colorScheme.error.copy(alpha = 0.5f)
            else MaterialTheme.colorScheme.outlineVariant.copy(alpha = 0.5f)),
        modifier = Modifier.fillMaxWidth().clickable(onClick = onClick)
    ) {
        Row(Modifier.padding(horizontal = 14.dp, vertical = 12.dp),
            verticalAlignment = Alignment.CenterVertically) {
            Column(Modifier.weight(1f)) {
                Text(field, style = MaterialTheme.typography.titleSmall)
                Spacer(Modifier.height(5.dp))
                when {
                    config != null && config.source == "column" ->
                        StatusChip(config.chipText(),
                            PosterTheme.extended.successContainer,
                            PosterTheme.extended.onSuccessContainer)
                    config != null && config.source == "fixed" ->
                        StatusChip(config.chipText(),
                            MaterialTheme.colorScheme.secondaryContainer,
                            MaterialTheme.colorScheme.onSecondaryContainer)
                    config != null -> StatusChip("自动",
                        MaterialTheme.colorScheme.surfaceContainerHigh,
                        MaterialTheme.colorScheme.onSurfaceVariant)
                    autoMatched != null ->
                        StatusChip("列『$autoMatched』",
                            PosterTheme.extended.successContainer,
                            PosterTheme.extended.onSuccessContainer)
                    else -> StatusChip("未绑定",
                        MaterialTheme.colorScheme.errorContainer,
                        MaterialTheme.colorScheme.onErrorContainer)
                }
            }
            Icon(Icons.Outlined.ChevronRight, contentDescription = null,
                tint = MaterialTheme.colorScheme.outline)
        }
    }
}

@Composable
private fun StatusChip(text: String, bg: androidx.compose.ui.graphics.Color, fg: androidx.compose.ui.graphics.Color) {
    Surface(shape = RoundedCornerShape(50), color = bg) {
        Text(text, modifier = Modifier.padding(horizontal = 10.dp, vertical = 3.dp),
            style = MaterialTheme.typography.labelSmall, color = fg,
            maxLines = 1, overflow = TextOverflow.Ellipsis)
    }
}

/** 字段配置底部抽屉：数据源（自动/数据列/固定值）+ 前后缀 */
@OptIn(ExperimentalMaterial3Api::class)
@Composable
private fun FieldConfigSheet(
    field: String,
    headers: List<String>,
    autoMatched: String?,
    initial: FieldConfig?,
    onSave: (FieldConfig?) -> Unit,
    onDismiss: () -> Unit
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

    Column(Modifier.padding(horizontal = 20.dp)) {
        Text(field, style = MaterialTheme.typography.titleLarge)
        Text(
            if (autoMatched != null) "自动匹配 → $autoMatched" else "未自动匹配",
            style = MaterialTheme.typography.bodySmall,
            color = MaterialTheme.colorScheme.onSurfaceVariant)
        Spacer(Modifier.height(14.dp))

        val options = listOf("自动匹配" to "auto", "数据列" to "column", "固定值" to "fixed")
        SingleChoiceSegmentedButtonRow(Modifier.fillMaxWidth()) {
            options.forEachIndexed { i, (label, key) ->
                SegmentedButton(
                    selected = source == key,
                    onClick = { source = key },
                    shape = SegmentedButtonDefaults.itemShape(index = i, count = options.size)
                ) { Text(label, maxLines = 1, overflow = TextOverflow.Ellipsis) }
            }
        }
        Spacer(Modifier.height(12.dp))

        when (source) {
            "column" -> {
                if (headers.isEmpty()) {
                    Text("（数据未加载，无列可选）",
                        style = MaterialTheme.typography.bodySmall,
                        color = MaterialTheme.colorScheme.onSurfaceVariant)
                } else {
                    Column(
                        Modifier.heightIn(max = 240.dp)
                            .verticalScroll(rememberScrollState())
                    ) {
                        headers.forEach { h ->
                            Row(
                                Modifier.fillMaxWidth()
                                    .clickable { column = h }
                                    .padding(vertical = 4.dp),
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                RadioButton(selected = h == column, onClick = { column = h })
                                Text(h, style = MaterialTheme.typography.bodyLarge)
                            }
                        }
                    }
                }
            }
            "fixed" -> {
                OutlinedTextField(
                    value = fixed, onValueChange = { fixed = it },
                    label = { Text("固定值（应用到所有行）") },
                    singleLine = true, modifier = Modifier.fillMaxWidth())
            }
        }
        Spacer(Modifier.height(12.dp))

        Row(horizontalArrangement = Arrangement.spacedBy(12.dp)) {
            OutlinedTextField(
                value = prefix, onValueChange = { prefix = it },
                label = { Text("前缀（可选）") },
                placeholder = { Text("如 ¥") },
                singleLine = true, modifier = Modifier.weight(1f))
            OutlinedTextField(
                value = suffix, onValueChange = { suffix = it },
                label = { Text("后缀（可选）") },
                placeholder = { Text("如 元") },
                singleLine = true, modifier = Modifier.weight(1f))
        }
        Spacer(Modifier.height(18.dp))
        Button(
            onClick = {
                when {
                    source == "auto" -> onSave(null)
                    source == "column" && column.isNotEmpty() ->
                        onSave(FieldConfig("column", column, prefix, suffix))
                    source == "fixed" -> onSave(FieldConfig("fixed", fixed, prefix, suffix))
                    else -> onDismiss()
                }
            },
            modifier = Modifier.fillMaxWidth().height(50.dp),
            shape = RoundedCornerShape(50)
        ) { Text("保存") }
        TextButton(
            onClick = onDismiss,
            modifier = Modifier.fillMaxWidth(),
        ) { Text("取消") }
        Spacer(Modifier.height(18.dp))
    }
}

@Composable
private fun EmptyState(text: String) {
    Box(Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
        Text(text, color = MaterialTheme.colorScheme.onSurfaceVariant)
    }
}
