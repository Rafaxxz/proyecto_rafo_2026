package pe.rafa.analizadorseace

import android.annotation.SuppressLint
import android.content.ActivityNotFoundException
import android.content.ContentValues
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.os.Environment
import android.os.Handler
import android.os.Looper
import android.provider.MediaStore
import android.provider.OpenableColumns
import android.util.Base64
import android.webkit.JavascriptInterface
import android.webkit.ValueCallback
import android.webkit.WebChromeClient
import android.webkit.WebResourceError
import android.webkit.WebResourceRequest
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Toast
import androidx.activity.OnBackPressedCallback
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import com.chaquo.python.Python
import com.chaquo.python.android.AndroidPlatform
import java.io.File

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private var fileCallback: ValueCallback<Array<Uri>>? = null
    private var reintentos = 0
    private val urlBase = "http://127.0.0.1:5000/"

    private val selectorWeb = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        val uris = urisDe(result.resultCode, result.data)
        fileCallback?.onReceiveValue(if (uris.isEmpty()) null else uris.toTypedArray())
        fileCallback = null
    }

    private val selectorNativo = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        val uris = urisDe(result.resultCode, result.data)
        if (uris.isEmpty()) {
            Toast.makeText(this, "No seleccionaste nada", Toast.LENGTH_SHORT).show()
            return@registerForActivityResult
        }
        val encolados = encolar(uris)
        if (encolados > 0) {
            webView.evaluateJavascript(
                "window.pendientesListos && window.pendientesListos()", null
            )
        }
    }

    @SuppressLint("SetJavaScriptEnabled")
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        if (!Python.isStarted()) {
            Python.start(AndroidPlatform(this))
        }
        Python.getInstance().getModule("servidor").callAttr("iniciar")

        webView = WebView(this)
        setContentView(webView)
        webView.settings.javaScriptEnabled = true
        webView.settings.domStorageEnabled = true
        webView.settings.allowFileAccess = true
        webView.settings.allowContentAccess = true
        webView.settings.javaScriptCanOpenWindowsAutomatically = true
        webView.addJavascriptInterface(Puente(), "AndroidBridge")

        webView.webViewClient = object : WebViewClient() {
            override fun onReceivedError(
                view: WebView?, request: WebResourceRequest?, error: WebResourceError?
            ) {
                if (request?.isForMainFrame == true && reintentos < 20) {
                    reintentos++
                    Handler(Looper.getMainLooper()).postDelayed({ webView.loadUrl(urlBase) }, 500)
                }
            }
        }

        webView.webChromeClient = object : WebChromeClient() {
            override fun onShowFileChooser(
                view: WebView?, callback: ValueCallback<Array<Uri>>?,
                params: FileChooserParams?
            ): Boolean {
                fileCallback?.onReceiveValue(null)
                fileCallback = callback
                return abrirSelector(selectorWeb) {
                    fileCallback?.onReceiveValue(null)
                    fileCallback = null
                }
            }
        }

        onBackPressedDispatcher.addCallback(this, object : OnBackPressedCallback(true) {
            override fun handleOnBackPressed() {
                if (webView.canGoBack()) {
                    webView.goBack()
                } else if (webView.url?.endsWith("/") == false) {
                    webView.loadUrl(urlBase)
                } else {
                    finish()
                }
            }
        })

        manejarIntent(intent)
        Handler(Looper.getMainLooper()).postDelayed({ webView.loadUrl(urlBase) }, 700)
    }

    private fun abrirSelector(
        launcher: androidx.activity.result.ActivityResultLauncher<Intent>,
        alFallar: () -> Unit
    ): Boolean {
        val abrir = Intent(Intent.ACTION_OPEN_DOCUMENT).apply {
            addCategory(Intent.CATEGORY_OPENABLE)
            type = "*/*"
            putExtra(Intent.EXTRA_ALLOW_MULTIPLE, true)
            addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        }
        try {
            launcher.launch(abrir)
            return true
        } catch (e: ActivityNotFoundException) {
            val alterno = Intent(Intent.ACTION_GET_CONTENT).apply {
                addCategory(Intent.CATEGORY_OPENABLE)
                type = "*/*"
                putExtra(Intent.EXTRA_ALLOW_MULTIPLE, true)
            }
            return try {
                launcher.launch(Intent.createChooser(alterno, "Selecciona PDFs"))
                true
            } catch (e2: Exception) {
                Toast.makeText(this, "No hay app de archivos disponible", Toast.LENGTH_LONG).show()
                alFallar()
                false
            }
        } catch (e: Exception) {
            Toast.makeText(this, "No pude abrir el selector: ${e.message}", Toast.LENGTH_LONG).show()
            alFallar()
            return false
        }
    }

    private fun urisDe(codigo: Int, datos: Intent?): List<Uri> {
        val uris = mutableListOf<Uri>()
        if (codigo != RESULT_OK || datos == null) return uris
        datos.clipData?.let { clip ->
            for (i in 0 until clip.itemCount) uris.add(clip.getItemAt(i).uri)
        }
        datos.data?.let { if (uris.isEmpty()) uris.add(it) }
        return uris
    }

    private fun encolar(uris: List<Uri>): Int {
        val servidor = Python.getInstance().getModule("servidor")
        val inbox = File(filesDir, "inbox").apply { mkdirs() }
        var encolados = 0
        uris.forEach { uri ->
            try {
                val nombre = nombreDe(uri)
                val destino = File(inbox, "${System.currentTimeMillis()}_$nombre")
                contentResolver.openInputStream(uri)?.use { entrada ->
                    destino.outputStream().use { salida -> entrada.copyTo(salida) }
                }
                servidor.callAttr("encolar", destino.absolutePath)
                encolados++
            } catch (e: Exception) {
                Toast.makeText(this, "No pude leer un PDF: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
        if (encolados > 0) {
            Toast.makeText(this, "$encolados PDF(s) recibido(s)", Toast.LENGTH_SHORT).show()
        }
        return encolados
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        setIntent(intent)
        if (manejarIntent(intent)) {
            webView.loadUrl(urlBase)
        }
    }

    private fun manejarIntent(intent: Intent?): Boolean {
        if (intent == null) return false
        val uris = mutableListOf<Uri>()
        when (intent.action) {
            Intent.ACTION_VIEW -> intent.data?.let { uris.add(it) }
            Intent.ACTION_SEND -> {
                @Suppress("DEPRECATION")
                (intent.getParcelableExtra(Intent.EXTRA_STREAM) as? Uri)?.let { uris.add(it) }
            }
            Intent.ACTION_SEND_MULTIPLE -> {
                @Suppress("DEPRECATION")
                intent.getParcelableArrayListExtra<Uri>(Intent.EXTRA_STREAM)?.let { uris.addAll(it) }
            }
        }
        if (uris.isEmpty()) return false
        return encolar(uris) > 0
    }

    private fun nombreDe(uri: Uri): String {
        var nombre = "documento.pdf"
        try {
            contentResolver.query(uri, null, null, null, null)?.use { c ->
                val idx = c.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                if (idx >= 0 && c.moveToFirst()) nombre = c.getString(idx)
            }
        } catch (e: Exception) {
            uri.lastPathSegment?.let { nombre = it }
        }
        if (!nombre.lowercase().endsWith(".pdf")) nombre += ".pdf"
        return nombre.replace(Regex("[^A-Za-z0-9._-]"), "_")
    }

    inner class Puente {
        @JavascriptInterface
        fun elegirArchivos() {
            runOnUiThread { abrirSelector(selectorNativo) {} }
        }

        @JavascriptInterface
        fun guardarArchivo(nombre: String, base64: String) {
            try {
                val datos = Base64.decode(base64, Base64.DEFAULT)
                val guardado = guardarEnDescargas(nombre, datos)
                runOnUiThread {
                    Toast.makeText(this@MainActivity, "Guardado en Descargas: $guardado", Toast.LENGTH_LONG).show()
                }
            } catch (e: Exception) {
                runOnUiThread {
                    Toast.makeText(this@MainActivity, "Error guardando: ${e.message}", Toast.LENGTH_LONG).show()
                }
            }
        }
    }

    /**
     * Guarda en Descargas reintentando con sufijos. Si en la carpeta ya quedo un
     * archivo con ese nombre de una instalacion anterior, MediaStore responde
     * "Failed to build unique file" en vez de renombrarlo solo; el sufijo esquiva eso.
     */
    private fun guardarEnDescargas(nombre: String, datos: ByteArray): String {
        val limpio = nombre.replace(Regex("[\\\\/:*?\"<>|]"), "_").ifBlank { "archivo.bin" }
        var ultimoError: Exception? = null
        for (intento in 0..4) {
            val candidato = if (intento == 0) limpio else conSufijo(limpio, intento)
            var uri: Uri? = null
            try {
                val valores = ContentValues().apply {
                    put(MediaStore.Downloads.DISPLAY_NAME, candidato)
                    put(MediaStore.Downloads.MIME_TYPE, mimeDe(candidato))
                    put(MediaStore.Downloads.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS)
                }
                uri = contentResolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, valores)
                    ?: throw java.io.IOException("Descargas rechazo el nombre")
                contentResolver.openOutputStream(uri)?.use { salida -> salida.write(datos) }
                    ?: throw java.io.IOException("No pude abrir el archivo para escribir")
                return candidato
            } catch (e: Exception) {
                ultimoError = e
                // Sin esto quedaria un archivo vacio ocupando el nombre.
                uri?.let { try { contentResolver.delete(it, null, null) } catch (_: Exception) {} }
            }
        }
        throw ultimoError ?: java.io.IOException("No pude guardar $limpio")
    }

    private fun conSufijo(nombre: String, n: Int): String {
        val punto = nombre.lastIndexOf('.')
        return if (punto > 0) "${nombre.substring(0, punto)}_$n${nombre.substring(punto)}"
        else "${nombre}_$n"
    }

    private fun mimeDe(nombre: String): String = when {
        nombre.endsWith(".pdf", true) -> "application/pdf"
        nombre.endsWith(".xlsx", true) ->
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else -> "application/octet-stream"
    }
}
