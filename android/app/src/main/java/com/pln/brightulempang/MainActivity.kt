package com.pln.brightulempang

import android.annotation.SuppressLint
import android.app.Activity
import android.content.ContentValues
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.provider.MediaStore
import android.util.Base64
import android.webkit.JavascriptInterface
import android.webkit.WebChromeClient
import android.webkit.WebResourceRequest
import android.webkit.WebSettings
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Toast
import java.io.File
import java.io.FileOutputStream

class MainActivity : Activity() {
 private lateinit var web: WebView
 private val dashboardUrl = "https://testproject966.github.io/dashboardulpempang/?app=android&mobile=1&v=5"
 inner class AndroidBridge {
  @JavascriptInterface fun savePpt(base64:String,fileName:String){try{
   val bytes=Base64.decode(base64,Base64.DEFAULT)
   if(Build.VERSION.SDK_INT>=Build.VERSION_CODES.Q){
    val values=ContentValues().apply{put(MediaStore.Downloads.DISPLAY_NAME,fileName);put(MediaStore.Downloads.MIME_TYPE,"application/vnd.openxmlformats-officedocument.presentationml.presentation");put(MediaStore.Downloads.RELATIVE_PATH,Environment.DIRECTORY_DOWNLOADS);put(MediaStore.Downloads.IS_PENDING,1)}
    val uri:Uri=contentResolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI,values)?:throw Exception("Tidak dapat membuat file Download")
    contentResolver.openOutputStream(uri).use{out->out?:throw Exception("Tidak dapat membuka file");out.write(bytes)}
    values.clear();values.put(MediaStore.Downloads.IS_PENDING,0);contentResolver.update(uri,values,null,null)
   }else{@Suppress("DEPRECATION") val dir=Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS);if(!dir.exists())dir.mkdirs();FileOutputStream(File(dir,fileName)).use{it.write(bytes)}}
   runOnUiThread{Toast.makeText(this@MainActivity,"PPT tersimpan di folder Download",Toast.LENGTH_LONG).show()}
  }catch(e:Exception){runOnUiThread{Toast.makeText(this@MainActivity,"Gagal menyimpan PPT: ${e.message?:"error"}",Toast.LENGTH_LONG).show()}}}
 }
 @SuppressLint("SetJavaScriptEnabled") override fun onCreate(savedInstanceState:Bundle?){super.onCreate(savedInstanceState);web=WebView(this);setContentView(web)
  web.settings.javaScriptEnabled=true;web.settings.domStorageEnabled=true;web.settings.databaseEnabled=true;web.settings.loadsImagesAutomatically=true;web.settings.allowFileAccess=false;web.settings.javaScriptCanOpenWindowsAutomatically=true;web.settings.setSupportZoom(false);web.settings.builtInZoomControls=false;web.settings.displayZoomControls=false;web.settings.cacheMode=WebSettings.LOAD_NO_CACHE
  web.overScrollMode=WebView.OVER_SCROLL_NEVER;web.isVerticalScrollBarEnabled=false;web.isHorizontalScrollBarEnabled=false;web.addJavascriptInterface(AndroidBridge(),"AndroidBridge");web.webChromeClient=WebChromeClient();web.webViewClient=object:WebViewClient(){override fun shouldOverrideUrlLoading(view:WebView,request:WebResourceRequest)=false};web.loadUrl(dashboardUrl+"&_="+System.currentTimeMillis())
 }
 @Deprecated("Deprecated in Android API 33") override fun onBackPressed(){if(web.canGoBack())web.goBack()else super.onBackPressed()}
 override fun onDestroy(){web.stopLoading();web.removeJavascriptInterface("AndroidBridge");web.destroy();super.onDestroy()}
}
