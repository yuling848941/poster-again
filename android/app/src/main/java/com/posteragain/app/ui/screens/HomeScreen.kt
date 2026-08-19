package com.posteragain.app.ui.screens

import android.graphics.Bitmap
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.Image
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.clickable
import androidx.compose.foundation.horizontalScroll
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
import androidx.compose.foundation.shape.CircleShape
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.Add
import androidx.compose.material.icons.outlined.AutoAwesome
import androidx.compose.material.icons.outlined.CheckCircle
import androidx.compose.material.icons.outlined.Description
import androidx.compose.material.icons.outlined.History
import androidx.compose.material.icons.outlined.TableChart
import androidx.compose.material3.Button
import androidx.compose.material3.ButtonDefaults
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.Icon
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Surface
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.remember
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.asImageBitmap
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import com.posteragain.app.AppState
import com.posteragain.app.Route
import com.posteragain.app.decodeScaled
import com.posteragain.app.ui.theme.PosterTheme
import java.io.File

/** 首页：三步向导（选模板 → 选数据 → 开始生成）+ 最近使用模板 */
@Composable
fun HomeScreen(state: AppState) {
    val tplPicker = rememberLauncherForActivityResult(ActivityResultContracts.GetContent()) { uri ->
        uri?.let { state.pickUri(it, isTemplate = true) }
    }
    val dataPicker = rememberLauncherForActivityResult(ActivityResultContracts.GetContent()) { uri ->
        uri?.let { state.pickUri(it, isTemplate = false) }
    }

    Column(
        Modifier
            .fillMaxSize()
            .verticalScroll(rememberScrollState())
            .padding(horizontal = 16.dp)
    ) {
        // 顶栏（按设计稿：居中 App 名，设置入口只在底部导航，避免重复）
        Box(Modifier.fillMaxWidth().height(64.dp), contentAlignment = Alignment.Center) {
            Text("PosterAgain",
                style = MaterialTheme.typography.titleLarge,
                fontWeight = FontWeight.Bold,
                color = MaterialTheme.colorScheme.primary)
        }
        // 意图说明
        Column(
            Modifier.fillMaxWidth().padding(vertical = 4.dp),
            horizontalAlignment = Alignment.CenterHorizontally
        ) {
            Text("批量生成个性化海报", style = MaterialTheme.typography.titleLarge)
            Text("快速连接数据与模板，一键生成",
                style = MaterialTheme.typography.bodyMedium,
                color = MaterialTheme.colorScheme.onSurfaceVariant)
        }

        // ---- 三步向导卡片 ----
        Surface(
            modifier = Modifier.fillMaxWidth().padding(top = 16.dp),
            shape = RoundedCornerShape(24.dp),
            color = MaterialTheme.colorScheme.surfaceContainerLow,
            border = androidx.compose.foundation.BorderStroke(
                1.dp, MaterialTheme.colorScheme.outlineVariant.copy(alpha = 0.4f)),
            shadowElevation = 2.dp
        ) {
            Column(Modifier.padding(20.dp), verticalArrangement = Arrangement.spacedBy(16.dp)) {
                // 第 1 步：模板
                StepRow(number = 1, connector = true, title = "选择模板",
                    done = state.tplBytes != null && state.session != null) {
                    if (state.tplBytes != null) {
                        SelectedTemplateCard(
                            name = state.tplName,
                            thumb = state.tplHash?.let {
                                state.thumbFile(it).takeIf { f -> f.exists() }?.let { f ->
                                    remember(f.path) { decodeScaled(f.path, 256) }
                                }
                            },
                            sub = state.session?.let { "${it.fields.size} 个占位符" }
                                ?: "${(state.tplBytes?.size ?: 0) / 1024} KB",
                            onChange = { tplPicker.launch("*/*") })
                    } else {
                        EmptyPickBox("点击选择 PPT 模板文件（.pptx）") { tplPicker.launch("*/*") }
                    }
                }
                // 第 2 步：数据
                StepRow(number = 2, connector = false, title = "选择数据",
                    done = state.dataBytes != null) {
                    if (state.dataBytes != null) {
                        SelectedDataCard(
                            name = state.dataName,
                            rowCount = state.session?.rowCount ?: 0,
                            onChange = { dataPicker.launch("*/*") })
                    } else {
                        EmptyPickBox("点击选择数据文件（Excel / CSV / JSON）") { dataPicker.launch("*/*") }
                    }
                }
                // 第 3 步：开始生成
                val s = state.session
                val ready = s != null && s.rowCount > 0 && !state.loadingSession &&
                        state.phase != AppState.Phase.Rendering
                Column(horizontalAlignment = Alignment.CenterHorizontally) {
                    Button(
                        onClick = { state.navigate(Route.Binding) },
                        enabled = ready,
                        modifier = Modifier.fillMaxWidth().height(52.dp),
                        shape = RoundedCornerShape(50)
                    ) {
                        Icon(Icons.Outlined.AutoAwesome, contentDescription = null,
                            modifier = Modifier.size(18.dp))
                        Spacer(Modifier.width(8.dp))
                        Text(if (s != null && s.rowCount > 0) "开始生成（${s.rowCount} 张）" else "开始生成")
                    }
                    if (state.loadingSession) {
                        Spacer(Modifier.height(10.dp))
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            CircularProgressIndicator(
                                strokeWidth = 2.dp, modifier = Modifier.size(14.dp))
                            Spacer(Modifier.width(8.dp))
                            Text("正在解析模板…", style = MaterialTheme.typography.bodySmall,
                                color = MaterialTheme.colorScheme.onSurfaceVariant)
                        }
                    } else if (s == null && state.tplBytes != null) {
                        Spacer(Modifier.height(6.dp))
                        Text("模板解析失败或未完成，请重试",
                            style = MaterialTheme.typography.bodySmall,
                            color = MaterialTheme.colorScheme.error)
                    } else if (s != null && s.rowCount == 0) {
                        Spacer(Modifier.height(6.dp))
                        Text("模板就绪，还需选择数据文件",
                            style = MaterialTheme.typography.bodySmall,
                            color = MaterialTheme.colorScheme.onSurfaceVariant)
                    }
                }
            }
        }

        // ---- 最近使用 ----
        if (state.recents.isNotEmpty()) {
            Text("最近使用", style = MaterialTheme.typography.titleLarge,
                modifier = Modifier.padding(start = 4.dp, top = 24.dp, bottom = 8.dp))
            Row(
                Modifier.fillMaxWidth().horizontalScroll(rememberScrollState()),
                horizontalArrangement = Arrangement.spacedBy(12.dp)
            ) {
                state.recents.forEach { r ->
                    RecentTemplateCard(
                        name = r.name,
                        count = r.count,
                        thumb = state.thumbFile(r.hash).takeIf { it.exists() }
                            ?.let { f -> remember(f.path) { decodeScaled(f.path, 256) } },
                        onClick = { state.loadRecentTemplate(r) })
                }
                // 「浏览更多模板」入口
                Column(Modifier.width(108.dp)) {
                    Surface(
                        modifier = Modifier.fillMaxWidth().height(144.dp),
                        shape = RoundedCornerShape(12.dp),
                        color = MaterialTheme.colorScheme.surfaceContainerLowest,
                        border = androidx.compose.foundation.BorderStroke(
                            1.dp, MaterialTheme.colorScheme.outlineVariant.copy(alpha = 0.4f))
                    ) {
                        Box(Modifier.clickable { state.navigate(Route.Templates) },
                            contentAlignment = Alignment.Center) {
                            Icon(Icons.Outlined.Add, contentDescription = null,
                                tint = MaterialTheme.colorScheme.outline,
                                modifier = Modifier.size(30.dp))
                        }
                    }
                    Spacer(Modifier.height(6.dp))
                    Text("浏览更多模板", style = MaterialTheme.typography.bodySmall,
                        color = MaterialTheme.colorScheme.onSurfaceVariant)
                }
            }
        }
        Spacer(Modifier.height(24.dp))
    }
}

@Composable
private fun StepRow(
    number: Int,
    connector: Boolean,
    title: String,
    done: Boolean,
    content: @Composable () -> Unit
) {
    Row(Modifier.fillMaxWidth()) {
        Column(horizontalAlignment = Alignment.CenterHorizontally) {
            Box(
                Modifier.size(30.dp).background(
                    MaterialTheme.colorScheme.primary, CircleShape),
                contentAlignment = Alignment.Center
            ) {
                Text("$number", color = MaterialTheme.colorScheme.onPrimary,
                    style = MaterialTheme.typography.labelLarge)
            }
            if (connector) {
                Spacer(Modifier.height(4.dp))
                Box(Modifier.width(2.dp).height(44.dp)
                    .background(MaterialTheme.colorScheme.outlineVariant.copy(alpha = 0.4f)))
            }
        }
        Column(Modifier.weight(1f).padding(start = 14.dp)) {
            Row(verticalAlignment = Alignment.CenterVertically) {
                Text(title, style = MaterialTheme.typography.labelLarge)
                if (done) {
                    Spacer(Modifier.width(6.dp))
                    Icon(Icons.Outlined.CheckCircle, contentDescription = null,
                        tint = PosterTheme.extended.success, modifier = Modifier.size(16.dp))
                }
            }
            Spacer(Modifier.height(8.dp))
            content()
        }
    }
}

@Composable
private fun EmptyPickBox(text: String, onClick: () -> Unit) {
    Surface(
        modifier = Modifier.fillMaxWidth().clickable(onClick = onClick),
        shape = RoundedCornerShape(12.dp),
        color = MaterialTheme.colorScheme.surfaceContainerLowest,
        border = androidx.compose.foundation.BorderStroke(
            1.dp, MaterialTheme.colorScheme.outlineVariant)
    ) {
        Text(text,
            modifier = Modifier.padding(vertical = 26.dp),
            textAlign = TextAlign.Center,
            style = MaterialTheme.typography.bodyMedium,
            color = MaterialTheme.colorScheme.onSurfaceVariant)
    }
}

@Composable
private fun SelectedTemplateCard(
    name: String, thumb: Bitmap?, sub: String, onChange: () -> Unit
) {
    Surface(
        shape = RoundedCornerShape(12.dp),
        color = MaterialTheme.colorScheme.surfaceContainerLowest,
        border = androidx.compose.foundation.BorderStroke(
            1.dp, MaterialTheme.colorScheme.outlineVariant.copy(alpha = 0.6f))
    ) {
        Row(Modifier.padding(10.dp), verticalAlignment = Alignment.CenterVertically) {
            if (thumb != null) {
                Image(thumb.asImageBitmap(), contentDescription = null,
                    modifier = Modifier.width(46.dp).height(60.dp)
                        .border(1.dp, MaterialTheme.colorScheme.outlineVariant.copy(alpha = 0.4f),
                            RoundedCornerShape(6.dp)),
                    contentScale = androidx.compose.ui.layout.ContentScale.Crop)
            } else {
                Box(Modifier.width(46.dp).height(60.dp)
                    .background(MaterialTheme.colorScheme.surfaceContainerHigh,
                        RoundedCornerShape(6.dp)),
                    contentAlignment = Alignment.Center) {
                    Icon(Icons.Outlined.Description, contentDescription = null,
                        tint = MaterialTheme.colorScheme.outline, modifier = Modifier.size(20.dp))
                }
            }
            Column(Modifier.weight(1f).padding(horizontal = 12.dp)) {
                Text(name, style = MaterialTheme.typography.bodyMedium,
                    maxLines = 1, overflow = TextOverflow.Ellipsis)
                Text(sub, style = MaterialTheme.typography.labelSmall,
                    color = MaterialTheme.colorScheme.onSurfaceVariant)
            }
            TextButton(onClick = onChange, contentPadding = androidx.compose.foundation.layout.PaddingValues(horizontal = 8.dp)) {
                Text("更换")
            }
        }
    }
}

@Composable
private fun SelectedDataCard(name: String, rowCount: Int, onChange: () -> Unit) {
    Surface(
        shape = RoundedCornerShape(12.dp),
        color = MaterialTheme.colorScheme.surfaceContainerLowest,
        border = androidx.compose.foundation.BorderStroke(
            1.dp, MaterialTheme.colorScheme.outlineVariant.copy(alpha = 0.6f))
    ) {
        Row(Modifier.padding(10.dp), verticalAlignment = Alignment.CenterVertically) {
            Box(Modifier.size(44.dp).background(
                PosterTheme.extended.successContainer, CircleShape),
                contentAlignment = Alignment.Center) {
                Icon(Icons.Outlined.TableChart, contentDescription = null,
                    tint = PosterTheme.extended.success, modifier = Modifier.size(22.dp))
            }
            Column(Modifier.weight(1f).padding(horizontal = 12.dp)) {
                Text(name, style = MaterialTheme.typography.bodyMedium,
                    maxLines = 1, overflow = TextOverflow.Ellipsis)
                Surface(shape = RoundedCornerShape(50),
                    color = MaterialTheme.colorScheme.secondaryContainer) {
                    Text(
                        if (rowCount > 0) "$rowCount 行数据" else "待加载",
                        modifier = Modifier.padding(horizontal = 8.dp, vertical = 2.dp),
                        style = MaterialTheme.typography.labelSmall,
                        color = MaterialTheme.colorScheme.onSecondaryContainer)
                }
            }
            TextButton(onClick = onChange, contentPadding = androidx.compose.foundation.layout.PaddingValues(horizontal = 8.dp)) {
                Text("更换")
            }
        }
    }
}

@Composable
private fun RecentTemplateCard(
    name: String, count: Int, thumb: Bitmap?, onClick: () -> Unit
) {
    Column(Modifier.width(108.dp).clickable(onClick = onClick)) {
        Surface(
            modifier = Modifier.fillMaxWidth().height(144.dp),
            shape = RoundedCornerShape(12.dp),
            color = MaterialTheme.colorScheme.surfaceContainer,
            border = androidx.compose.foundation.BorderStroke(
                1.dp, MaterialTheme.colorScheme.outlineVariant.copy(alpha = 0.3f))
        ) {
            if (thumb != null) {
                Image(thumb.asImageBitmap(), contentDescription = null,
                    modifier = Modifier.fillMaxSize(),
                    contentScale = androidx.compose.ui.layout.ContentScale.Crop)
            } else {
                Box(Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                    Icon(Icons.Outlined.Description, contentDescription = null,
                        tint = MaterialTheme.colorScheme.outline)
                }
            }
        }
        Spacer(Modifier.height(6.dp))
        Text(name, style = MaterialTheme.typography.bodySmall,
            maxLines = 1, overflow = TextOverflow.Ellipsis)
        Row(verticalAlignment = Alignment.CenterVertically) {
            Icon(Icons.Outlined.History, contentDescription = null,
                modifier = Modifier.size(12.dp),
                tint = MaterialTheme.colorScheme.onSurfaceVariant)
            Spacer(Modifier.width(3.dp))
            Text("使用 $count 次", style = MaterialTheme.typography.labelSmall,
                color = MaterialTheme.colorScheme.onSurfaceVariant)
        }
    }
}
