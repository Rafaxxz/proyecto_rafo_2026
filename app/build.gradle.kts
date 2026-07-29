plugins {
    id("com.android.application")
    id("org.jetbrains.kotlin.android")
    id("com.chaquo.python")
}

android {
    namespace = "pe.rafa.analizadorseace"
    compileSdk = 34

    defaultConfig {
        applicationId = "pe.rafa.analizadorseace"
        minSdk = 29
        targetSdk = 34
        versionCode = 1
        versionName = "1.0"
        ndk {
            abiFilters += listOf("arm64-v8a")
        }
    }

    buildTypes {
        release {
            isMinifyEnabled = false
        }
    }
    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }
    kotlinOptions {
        jvmTarget = "17"
    }
}

chaquopy {
    defaultConfig {
        version = "3.12"
        pip {
            install("flask")
            install("pdfplumber==0.9.0")
            install("requests")
            install("beautifulsoup4")
            install("openpyxl")
            install("fpdf2")
            install("reportlab")
        }
        extractPackages("analizador")
    }
}

dependencies {
    implementation("androidx.appcompat:appcompat:1.7.0")
    implementation("androidx.activity:activity-ktx:1.9.0")
}
