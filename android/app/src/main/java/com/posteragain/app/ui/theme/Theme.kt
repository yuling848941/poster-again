package com.posteragain.app.ui.theme

import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material3.ColorScheme
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Shapes
import androidx.compose.material3.lightColorScheme
import androidx.compose.runtime.Composable
import androidx.compose.runtime.CompositionLocalProvider
import androidx.compose.runtime.staticCompositionLocalOf
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.unit.dp

/**
 * 主题：色板来自 Stitch 导出的设计系统（PosterAgain Design System）。
 * primary 为深青 #004D64（青色板与 DESIGN.md 保持一致，仅修正了导出稿中
 * 明显失衡的 on-primary-container 配对）。
 */

private val Primary = Color(0xFF004D64)
private val OnPrimary = Color(0xFFFFFFFF)
private val PrimaryContainer = Color(0xFFBEE9FF)
private val OnPrimaryContainer = Color(0xFF001F2A)

private val Secondary = Color(0xFF4D616C)
private val OnSecondary = Color(0xFFFFFFFF)
private val SecondaryContainer = Color(0xFFD0E6F3)
private val OnSecondaryContainer = Color(0xFF081E27)

private val Tertiary = Color(0xFF434365)
private val OnTertiary = Color(0xFFFFFFFF)
private val TertiaryContainer = Color(0xFFE1DFFF)
private val OnTertiaryContainer = Color(0xFF181837)

private val BgSurface = Color(0xFFF8FAFB)
private val OnBgSurface = Color(0xFF1B1C1C)
private val OnBgSurfaceVariant = Color(0xFF3F484D)

/** M3 标准 scheme 之外的语义色（成功态），用于「已匹配」chip 等 */
data class ExtendedColors(
    val success: Color = Color(0xFF006D39),
    val onSuccess: Color = Color(0xFFFFFFFF),
    val successContainer: Color = Color(0xFFD9F0E2),
    val onSuccessContainer: Color = Color(0xFF00532B),
)

val LocalExtendedColors = staticCompositionLocalOf { ExtendedColors() }

/** 便捷访问：PosterTheme.extended.success */
object PosterTheme {
    val extended: ExtendedColors
        @Composable get() = LocalExtendedColors.current

    val colors: ColorScheme
        @Composable get() = androidx.compose.material3.MaterialTheme.colorScheme
}

/** 深青主色的浅色 scheme（设计稿仅提供浅色模式，暂不跟随系统深色） */
val PosterAgainLightScheme = lightColorScheme(
    primary = Primary,
    onPrimary = OnPrimary,
    primaryContainer = PrimaryContainer,
    onPrimaryContainer = OnPrimaryContainer,
    inversePrimary = Color(0xFF87D0F2),
    secondary = Secondary,
    onSecondary = OnSecondary,
    secondaryContainer = SecondaryContainer,
    onSecondaryContainer = OnSecondaryContainer,
    tertiary = Tertiary,
    onTertiary = OnTertiary,
    tertiaryContainer = TertiaryContainer,
    onTertiaryContainer = OnTertiaryContainer,
    background = BgSurface,
    onBackground = OnBgSurface,
    surface = BgSurface,
    onSurface = OnBgSurface,
    surfaceVariant = Color(0xFFE4E2E2),
    onSurfaceVariant = OnBgSurfaceVariant,
    surfaceTint = Primary,
    inverseSurface = Color(0xFF303030),
    inverseOnSurface = Color(0xFFF2F0F0),
    outline = Color(0xFF70787E),
    outlineVariant = Color(0xFFBFC8CD),
    error = Color(0xFFBA1A1A),
    onError = Color(0xFFFFFFFF),
    errorContainer = Color(0xFFFFDAD6),
    onErrorContainer = Color(0xFF93000A),
    surfaceContainerLowest = Color(0xFFFFFFFF),
    surfaceContainerLow = Color(0xFFF0F4F6),
    surfaceContainer = Color(0xFFEFEDED),
    surfaceContainerHigh = Color(0xFFE9E8E7),
    surfaceContainerHighest = Color(0xFFE4E2E2),
    scrim = Color(0xFF000000),
)

/** 圆角体系：小件 8 / 卡片 16 / 大面 28（与设计系统一致） */
val PosterAgainShapes = Shapes(
    extraSmall = RoundedCornerShape(8.dp),
    small = RoundedCornerShape(8.dp),
    medium = RoundedCornerShape(12.dp),
    large = RoundedCornerShape(16.dp),
    extraLarge = RoundedCornerShape(28.dp),
)

@Composable
fun PosterAgainTheme(content: @Composable () -> Unit) {
    CompositionLocalProvider(LocalExtendedColors provides ExtendedColors()) {
        MaterialTheme(
            colorScheme = PosterAgainLightScheme,
            shapes = PosterAgainShapes,
            content = content
        )
    }
}
