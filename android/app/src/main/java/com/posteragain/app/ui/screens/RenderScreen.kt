package com.posteragain.app.ui.screens

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
import androidx.compose.foundation.lazy.grid.GridCells
import androidx.compose.foundation.lazy.grid.LazyVerticalGrid
import androidx.compose.foundation.lazy.grid.itemsIndexed
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.outlined.ArrowBack
import androidx.compose.material.icons.filled.CheckCircle
import androidx.compose.material.icons.outlined.Share
import androidx.compose.material3.Button
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.FilledTonalButton
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Surface
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.remember
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.layout.ContentScale
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.foundation.Image
import androidx.compose.foundation.clickable
import androidx.compose.ui.graphics.asImageBitmap
import com.posteragain.app.AppState
import com.posteragain.app.Route
import com.posteragain.app.decodeScaled
import com.posteragain.app.shareImages
import com.posteragain.app.ui.theme.PosterTheme
import java.io.File

/** 渲染页：进行中（进度环 + 已完成网格）→ 完成（网格 + 悬浮操作栏） */
@Composable
fun RenderScreen(state: AppState) {
    // 进入该屏即自动开始批量渲染
    LaunchedEffect(state.session?.sid) {
        if (state.phase == AppState.Phase.Idle && state.session?.rowCount ?: 0 > 0) {
            state.startBatch()
        }
    }

    Column(Modifier.fillMaxSize()) {
        Row(
            Modifier.fillMaxWidth().height(56.dp).padding(horizontal = 4.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            IconButton(onClick = { state.goBack() }) {
                Icon(Icons.AutoMirrored.Outlined.ArrowBack, contentDescription = "返回")
            }
            Text(
                if (state.phase == AppState.Phase.Rendering) "正在生成"
                else if (state.results.isEmpty()) "生成" else "生成完成",
                style = MaterialTheme.typography.titleLarge)
        }

        when (state.phase) {
            AppState.Phase.Rendering -> RenderingBody(state)
            else -> ResultBody(state)
        }
    }
}

@Composable
private fun RenderingBody(state: AppState) {
    val (done, total) = state.progress
    val remainingSec = if (state.avgMs > 0 && total > done)
        ((total - done) * state.avgMs / 1000).toInt() else null

    Column(
        Modifier.fillMaxSize().padding(horizontal = 24.dp),
        horizontalAlignment = Alignment.CenterHorizontally
    ) {
        Spacer(Modifier.height(32.dp))
        Box(contentAlignment = Alignment.Center) {
            CircularProgressIndicator(
                progress = { if (total > 0) done.toFloat() / total else 0f },
                modifier = Modifier.size(120.dp),
                strokeWidth = 8.dp,
                color = MaterialTheme.colorScheme.primary)
            Column(horizontalAlignment = Alignment.CenterHorizontally) {
                Text("$done", style = MaterialTheme.typography.headlineLarge,
                    fontWeight = FontWeight.Bold)
                Text("/ $total", style = MaterialTheme.typography.bodyMedium,
                    color = MaterialTheme.colorScheme.onSurfaceVariant)
            }
        }
        Spacer(Modifier.height(12.dp))
        Text(
            remainingSec?.let { "约 $it 秒" } ?: "正在渲染…",
            style = MaterialTheme.typography.bodyMedium,
            color = MaterialTheme.colorScheme.onSurfaceVariant)
        Spacer(Modifier.height(24.dp))
        // 已完成部分先以缩略图填充
        if (state.results.isNotEmpty()) {
            ResultGrid(state, Modifier.weight(1f))
        }
        TextButton(onClick = { state.cancelBatch() }, modifier = Modifier.padding(8.dp)) {
            Text("取消")
        }
        Spacer(Modifier.height(8.dp))
    }
}

@Composable
private fun ResultBody(state: AppState) {
    if (state.results.isEmpty()) {
        Box(Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
            Text("没有可展示的海报", color = MaterialTheme.colorScheme.onSurfaceVariant)
        }
        return
    }
    Column(Modifier.fillMaxSize()) {
        // 完成横幅
        Row(
            Modifier.fillMaxWidth().padding(horizontal = 20.dp, vertical = 8.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            Icon(Icons.Filled.CheckCircle,
                contentDescription = null,
                tint = PosterTheme.extended.success, modifier = Modifier.size(28.dp))
            Spacer(Modifier.width(10.dp))
            Text("${state.results.size} 张海报已就绪",
                style = MaterialTheme.typography.titleMedium)
        }
        ResultGrid(state, Modifier.weight(1f))
        FloatingActionBar(state)
    }
}

@Composable
private fun ResultGrid(state: AppState, modifier: Modifier = Modifier) {
    LazyVerticalGrid(
        columns = GridCells.Fixed(3),
        modifier = modifier.fillMaxWidth().padding(horizontal = 12.dp),
        horizontalArrangement = Arrangement.spacedBy(8.dp),
        verticalArrangement = Arrangement.spacedBy(8.dp)
    ) {
        itemsIndexed(state.results) { i, path ->
            ResultCell(path = path, onClick = { state.openPreview(i) })
        }
    }
}

@Composable
private fun ResultCell(path: String, onClick: () -> Unit) {
    val bmp = remember(path) { decodeScaled(path, 256) }
    Column(Modifier.clickable(onClick = onClick)) {
        Surface(
            shape = RoundedCornerShape(10.dp),
            color = MaterialTheme.colorScheme.surfaceContainer
        ) {
            if (bmp != null) {
                Image(bmp.asImageBitmap(), contentDescription = null,
                    modifier = Modifier.fillMaxWidth().height(160.dp),
                    contentScale = ContentScale.Crop)
            } else {
                Box(Modifier.fillMaxWidth().height(160.dp),
                    contentAlignment = Alignment.Center) {
                    Text("?", color = MaterialTheme.colorScheme.outline)
                }
            }
        }
        Text(
            captionOf(path),
            style = MaterialTheme.typography.labelSmall,
            maxLines = 1, overflow = TextOverflow.Ellipsis,
            textAlign = TextAlign.Center,
            modifier = Modifier.fillMaxWidth().padding(top = 3.dp))
    }
}

/** 由渲染产物文件名（如 1_陈小明.png）得到展示名 */
private fun captionOf(path: String): String {
    val n = File(path).nameWithoutExtension
    return n.substringAfter('_', n)
}

@Composable
private fun FloatingActionBar(state: AppState) {
    val context = LocalContext.current
    Surface(
        modifier = Modifier.fillMaxWidth().padding(12.dp),
        shape = RoundedCornerShape(50),
        color = MaterialTheme.colorScheme.surfaceContainerLowest,
        shadowElevation = 6.dp,
        border = androidx.compose.foundation.BorderStroke(
            1.dp, MaterialTheme.colorScheme.outlineVariant.copy(alpha = 0.4f))
    ) {
        Row(
            Modifier.padding(10.dp),
            horizontalArrangement = Arrangement.spacedBy(10.dp)
        ) {
            FilledTonalButton(
                onClick = { shareImages(context, state.results) },
                modifier = Modifier.weight(1f).height(46.dp),
                shape = RoundedCornerShape(50)
            ) {
                Icon(Icons.Outlined.Share, contentDescription = null,
                    modifier = Modifier.size(18.dp))
                Spacer(Modifier.width(6.dp))
                Text("分享")
            }
            Button(
                onClick = { state.saveAllToGallery() },
                modifier = Modifier.weight(1.4f).height(46.dp),
                shape = RoundedCornerShape(50)
            ) {
                Text("全部保存到相册")
            }
        }
    }
}
