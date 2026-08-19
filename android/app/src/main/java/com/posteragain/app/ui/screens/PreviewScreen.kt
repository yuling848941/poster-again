@file:OptIn(androidx.compose.foundation.ExperimentalFoundationApi::class)

package com.posteragain.app.ui.screens

import androidx.compose.foundation.Image
import androidx.compose.foundation.background
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
import androidx.compose.foundation.pager.HorizontalPager
import androidx.compose.foundation.pager.rememberPagerState
import androidx.compose.foundation.shape.CircleShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.outlined.ArrowBack
import androidx.compose.material.icons.outlined.Delete
import androidx.compose.material.icons.outlined.Download
import androidx.compose.material.icons.outlined.Share
import androidx.compose.material.icons.outlined.Wallpaper
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.remember
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Brush
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.asImageBitmap
import androidx.compose.ui.layout.ContentScale
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import com.posteragain.app.AppState
import com.posteragain.app.decodeScaled
import com.posteragain.app.saveImageToGallery
import com.posteragain.app.setAsWallpaper
import com.posteragain.app.shareImage
import java.io.File

/** 大图预览：左右滑动切换，底部四个操作（保存/分享/设为壁纸/删除） */
@Composable
fun PreviewScreen(state: AppState) {
    val context = LocalContext.current
    val paths = state.results
    if (paths.isEmpty()) {
        Box(Modifier.fillMaxSize().background(Color(0xFF101418)),
            contentAlignment = Alignment.Center) {
            Text("没有可预览的海报", color = Color(0xFFBFC8CD))
        }
        return
    }
    val pagerState = rememberPagerState(
        initialPage = state.previewIndex.coerceIn(0, paths.lastIndex)) {
        paths.size
    }

    Box(Modifier.fillMaxSize().background(Color(0xFF101418))) {
        HorizontalPager(state = pagerState, modifier = Modifier.fillMaxSize()) { page ->
            val path = paths[page]
            // 全屏查看，最长边控制在安全分辨率内
            val bmp = remember(path) { decodeScaled(path, 2200) }
            bmp?.let {
                Image(it.asImageBitmap(), contentDescription = null,
                    modifier = Modifier.fillMaxSize().padding(vertical = 76.dp, horizontal = 8.dp),
                    contentScale = ContentScale.Fit)
            }
        }

        // 顶部半透明栏：返回 + 文件名
        Box(
            Modifier.fillMaxWidth()
                .height(72.dp)
                .background(Brush.verticalGradient(
                    listOf(Color(0xB3000000), Color(0x00000000))))
                .align(Alignment.TopCenter)
        ) {
            Row(Modifier.padding(horizontal = 4.dp), verticalAlignment = Alignment.CenterVertically) {
                IconButton(onClick = { state.goBack() }) {
                    Icon(Icons.AutoMirrored.Outlined.ArrowBack,
                        contentDescription = "返回", tint = Color.White)
                }
                Text(File(paths[pagerState.currentPage]).name,
                    color = Color(0xFFE8EAED),
                    style = MaterialTheme.typography.bodyMedium,
                    maxLines = 1, overflow = TextOverflow.Ellipsis)
            }
        }

        // 底部半透明栏：页码 + 四个操作
        Box(
            Modifier.fillMaxWidth()
                .align(Alignment.BottomCenter)
                .background(Brush.verticalGradient(
                    listOf(Color(0x00000000), Color(0xB3000000))))
        ) {
            Column(Modifier.padding(bottom = 14.dp, top = 20.dp)) {
                Text("${pagerState.currentPage + 1} / ${paths.size}",
                    color = Color(0xFFE8EAED),
                    style = MaterialTheme.typography.labelMedium,
                    modifier = Modifier.fillMaxWidth(),
                    textAlign = TextAlign.Center)
                Spacer(Modifier.height(6.dp))
                val current = paths[pagerState.currentPage]
                Row(
                    Modifier.fillMaxWidth().padding(horizontal = 12.dp),
                    horizontalArrangement = Arrangement.SpaceEvenly
                ) {
                    PreviewAction(Icons.Outlined.Download, "保存到相册") {
                        val uri = runCatching {
                            saveImageToGallery(context, current,
                                state.imageFormat,
                                if (state.exportSize == "1080p") 1080
                                else if (state.exportSize == "2k") 2048 else null)
                        }.getOrNull()
                        state.message = if (uri != null) "已保存到相册" else "保存失败"
                    }
                    PreviewAction(Icons.Outlined.Share, "分享") {
                        runCatching { shareImage(context, current) }
                            .onFailure { state.message = "分享失败: ${it.message}" }
                    }
                    PreviewAction(Icons.Outlined.Wallpaper, "设为壁纸") {
                        runCatching { setAsWallpaper(context, current) }
                            .onFailure { state.message = "设置失败: ${it.message}" }
                    }
                    PreviewAction(Icons.Outlined.Delete, "删除", danger = true) {
                        state.deleteResult(pagerState.currentPage)
                    }
                }
            }
        }
    }
}

@Composable
private fun PreviewAction(
    icon: androidx.compose.ui.graphics.vector.ImageVector,
    label: String,
    danger: Boolean = false,
    onClick: () -> Unit
) {
    Column(
        horizontalAlignment = Alignment.CenterHorizontally,
        modifier = Modifier.padding(horizontal = 6.dp)
    ) {
        Box(
            Modifier.size(44.dp).background(Color(0x33FFFFFF), CircleShape),
            contentAlignment = Alignment.Center
        ) {
            IconButton(onClick = onClick, modifier = Modifier.size(40.dp)) {
                Icon(icon, contentDescription = label,
                    tint = if (danger) Color(0xFFFFB4AB) else Color(0xFFE8EAED),
                    modifier = Modifier.size(20.dp))
            }
        }
        Spacer(Modifier.height(3.dp))
        Text(label, color = Color(0xFFC9CDD2), style = MaterialTheme.typography.labelSmall)
    }
}
