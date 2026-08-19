package com.posteragain.app

import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.BackHandler
import androidx.activity.compose.setContent
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.padding
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.Assignment
import androidx.compose.material.icons.outlined.DashboardCustomize
import androidx.compose.material.icons.outlined.Settings
import androidx.compose.material3.Icon
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.NavigationBar
import androidx.compose.material3.NavigationBarItem
import androidx.compose.material3.Scaffold
import androidx.compose.material3.SnackbarHost
import androidx.compose.material3.SnackbarHostState
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.remember
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import com.chaquo.python.Python
import com.chaquo.python.android.AndroidPlatform
import com.posteragain.app.ui.screens.BindingScreen
import com.posteragain.app.ui.screens.HomeScreen
import com.posteragain.app.ui.screens.PreviewScreen
import com.posteragain.app.ui.screens.RenderScreen
import com.posteragain.app.ui.screens.SettingsScreen
import com.posteragain.app.ui.screens.TemplatesScreen
import com.posteragain.app.ui.theme.PosterAgainTheme

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        if (!Python.isStarted()) {
            Python.start(AndroidPlatform(this))
        }
        setContent {
            PosterAgainTheme {
                PosterAgainApp()
            }
        }
    }
}

/** 应用骨架：路由分发 + 底部导航（任务/模板/设置）+ 全局提示 */
@Composable
fun PosterAgainApp() {
    val context = LocalContext.current
    val state = remember { AppState(context.applicationContext) }
    val snackbarHostState = remember { SnackbarHostState() }

    // 全局一次性消息 → snackbar
    LaunchedEffect(state.message) {
        state.message?.let {
            snackbarHostState.showSnackbar(it)
            state.message = null
        }
    }
    BackHandler(enabled = state.canGoBack) { state.goBack() }

    Scaffold(
        containerColor = MaterialTheme.colorScheme.background,
        snackbarHost = { SnackbarHost(snackbarHostState) },
        bottomBar = {
            if (state.currentRoute.isTab) {
                NavigationBar {
                    NavigationBarItem(
                        selected = state.currentRoute == Route.Home,
                        onClick = { state.navigate(Route.Home) },
                        icon = { Icon(Icons.Outlined.Assignment, contentDescription = null) },
                        label = { Text("任务") })
                    NavigationBarItem(
                        selected = state.currentRoute == Route.Templates,
                        onClick = { state.navigate(Route.Templates) },
                        icon = { Icon(Icons.Outlined.DashboardCustomize, contentDescription = null) },
                        label = { Text("模板") })
                    NavigationBarItem(
                        selected = state.currentRoute == Route.Settings,
                        onClick = { state.navigate(Route.Settings) },
                        icon = { Icon(Icons.Outlined.Settings, contentDescription = null) },
                        label = { Text("设置") })
                }
            }
        }
    ) { padding ->
        Box(Modifier.fillMaxSize().padding(padding)) {
            when (state.currentRoute) {
                Route.Home -> HomeScreen(state)
                Route.Templates -> TemplatesScreen(state)
                Route.Settings -> SettingsScreen(state)
                Route.Binding -> BindingScreen(state)
                Route.Render -> RenderScreen(state)
                Route.Preview -> PreviewScreen(state)
            }
        }
    }
}
