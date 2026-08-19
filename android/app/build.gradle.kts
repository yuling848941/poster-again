import java.util.Properties

plugins {
    id("com.android.application")
    id("org.jetbrains.kotlin.android")
    id("com.chaquo.python")
}

// release 构建任务 → 单 ABI（arm64 真机分发，77MB）；
// 其他（assembleDebug/IDE 同步）→ 双 ABI（x86_64 供模拟器 E2E）。
// Chaquopy 要求 abiFilters 在 defaultConfig，buildType 级设置对其无效，
// 故按任务名在此切换（跑 "build" 生命周期任务时两变体都是双 ABI，安全但更大）。
val releaseOnlyBuild = gradle.startParameter.taskNames.any {
    it.contains("Release", ignoreCase = true)
}

android {
    namespace = "com.posteragain.app"
    compileSdk = 34

    defaultConfig {
        applicationId = "com.posteragain.app"
        minSdk = 24
        targetSdk = 34
        versionCode = 2
        versionName = "1.1.0"
        ndk {
            abiFilters += if (releaseOnlyBuild) listOf("arm64-v8a")
                          else listOf("arm64-v8a", "x86_64")
        }
    }

    buildFeatures {
        compose = true
        buildConfig = true   // 供设置页等读取 VERSION_NAME
    }
    composeOptions { kotlinCompilerExtensionVersion = "1.5.8" }

    signingConfigs {
        // 正式签名读 android/keystore.properties（gitignored，不入库）；
        // 文件缺失时回退 debug 签名，保证无证书环境也能构建。
        // ⚠ keystore + 口令必须离线备份，丢失则无法同签名覆盖升级！
        create("release") {
            val ksProps = Properties()
            val ksFile = rootProject.file("keystore.properties")
            if (ksFile.exists()) {
                ksFile.inputStream().use { ksProps.load(it) }
                storeFile = rootProject.file(ksProps.getProperty("storeFile"))
                storePassword = ksProps.getProperty("storePassword")
                keyAlias = ksProps.getProperty("keyAlias")
                keyPassword = ksProps.getProperty("keyPassword")
            } else {
                println("[签名警告] 未找到 keystore.properties，release 回退 debug 签名（仅限自用/测试）")
                storeFile = file(System.getProperty("user.home") + "/.android/debug.keystore")
                storePassword = "android"
                keyAlias = "androiddebugkey"
                keyPassword = "android"
            }
        }
    }

    buildTypes {
        release {
            signingConfig = signingConfigs.getByName("release")
            // Chaquopy 内嵌 Python 依赖未配 R8 规则，关闭 minify/混淆
            isMinifyEnabled = false
        }
    }
    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }
    kotlinOptions { jvmTarget = "17" }
    packagingOptions {
        resources.excludes += "/META-INF/{AL2.0,LGPL2.1}"
    }
    applicationVariants.all {
        val variant = this
        outputs.all {
            (this as? com.android.build.gradle.internal.api.BaseVariantOutputImpl)?.outputFileName =
                "PosterAgain-${variant.versionName}-${variant.buildType.name}.apk"
        }
    }
}

chaquopy {
    defaultConfig {
        version = "3.12"
        pip {
            // lxml/Pillow/numpy 走 Chaquopy 预编译仓；python-pptx/openpyxl 纯 Python 走 PyPI
            // （openpyxl 供 xlsx 数据读取；渐变文字的 numpy 缺失时另有纯 Pillow 兜底）
            install("python-pptx")
            install("Pillow")
            install("lxml")
            install("numpy")
            install("openpyxl")
        }
    }
}

dependencies {
    val composeBom = platform("androidx.compose:compose-bom:2024.02.01")
    implementation(composeBom)
    implementation("androidx.activity:activity-compose:1.8.2")
    implementation("androidx.core:core-ktx:1.12.0")
    implementation("androidx.compose.ui:ui")
    implementation("androidx.compose.material3:material3")
    implementation("androidx.compose.material:material-icons-extended")
    implementation("androidx.compose.ui:ui-tooling-preview")
    implementation("org.jetbrains.kotlinx:kotlinx-coroutines-android:1.7.3")
}
