package com.kxtt.energy;

import android.Manifest;
import android.app.Activity;
import android.app.AlertDialog;
import android.app.DownloadManager;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.webkit.CookieManager;
import android.webkit.DownloadListener;
import android.webkit.URLUtil;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.ProgressBar;
import android.widget.Toast;

import com.chaquo.python.PyObject;
import com.chaquo.python.Python;

import java.io.File;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;

public class MainActivity extends Activity {
    static final int REQ_FILE = 1001, REQ_NOTIF = 1002;
    WebView web;
    ProgressBar bar;
    ValueCallback<Uri[]> fileCb;
    boolean started;

    @Override
    protected void onCreate(Bundle b) {
        super.onCreate(b);
        getWindow().setStatusBarColor(0xFF1A56DB);
        bar = new ProgressBar(this);
        bar.setIndeterminate(true);
        web = new WebView(this);
        android.widget.FrameLayout root = new android.widget.FrameLayout(this);
        root.addView(web, new android.widget.FrameLayout.LayoutParams(-1, -1));
        root.addView(bar, new android.widget.FrameLayout.LayoutParams(80, 80, android.view.Gravity.CENTER));
        setContentView(root);
        bar.setVisibility(android.view.View.GONE);
        web.getSettings().setJavaScriptEnabled(true);
        web.getSettings().setDomStorageEnabled(true);
        web.getSettings().setDatabaseEnabled(true);
        web.getSettings().setSupportZoom(false);
        web.getSettings().setBuiltInZoomControls(false);
        web.getSettings().setCacheMode(android.webkit.WebSettings.LOAD_DEFAULT);
        CookieManager.getInstance().setAcceptCookie(true);
        web.setWebViewClient(new WebViewClient() {
            @Override
            public void onPageStarted(WebView v, String url, android.graphics.Bitmap f) { bar.setVisibility(android.view.View.VISIBLE); }
            @Override
            public void onPageFinished(WebView v, String url) { bar.setVisibility(android.view.View.GONE); }
            @Override
            public boolean shouldOverrideUrlLoading(WebView v, String url) {
                if (url != null && (url.startsWith("http://127.0.0.1") || url.startsWith("http://localhost"))) return false;
                if (url != null && (url.startsWith("http://") || url.startsWith("https://"))) {
                    startActivity(new Intent(Intent.ACTION_VIEW, Uri.parse(url)));
                    return true;
                }
                return false;
            }
        });
        web.setWebChromeClient(new WebChromeClient() {
            @Override
            public boolean onShowFileChooser(WebView v, ValueCallback<Uri[]> cb, FileChooserParams p) {
                if (fileCb != null) fileCb.onReceiveValue(null);
                fileCb = cb;
                try {
                    Intent i = new Intent(Intent.ACTION_GET_CONTENT);
                    i.addCategory(Intent.CATEGORY_OPENABLE);
                    i.setType("*/*");
                    i.putExtra(Intent.EXTRA_MIME_TYPES, new String[]{"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/octet-stream", "*/*"});
                    startActivityForResult(Intent.createChooser(i, "选择 Excel 文件"), REQ_FILE);
                } catch (Exception e) {
                    fileCb = null;
                    Toast.makeText(MainActivity.this, "没有找到文件管理器", Toast.LENGTH_SHORT).show();
                    return false;
                }
                return true;
            }
        });
        web.setDownloadListener(new DownloadListener() {
            @Override
            public void onDownloadStart(String url, String ua, String disp, String mime, long len) {
                save(url, disp, mime);
            }
        });
        startServer();
    }

    void startServer() {
        if (started) return;
        started = true;
        new Thread(new Runnable() {
            @Override
            public void run() {
                try {
                    if (!Python.isStarted()) Python.start(new com.chaquo.python.android.AndroidPlatform(MainActivity.this));
                    final int port = Python.getInstance().getModule("server.android_entry").callAttr("start", "127.0.0.1", 0).toInt();
                    runOnUiThread(new Runnable() {
                        @Override
                        public void run() { web.loadUrl("http://127.0.0.1:" + port + "/"); }
                    });
                } catch (final Exception e) {
                    runOnUiThread(new Runnable() {
                        @Override
                        public void run() { fatal("启动失败：\n" + e.getMessage()); }
                    });
                }
            }
        }).start();
    }

    void fatal(String msg) {
        new AlertDialog.Builder(this)
                .setTitle("能耗分析")
                .setMessage(msg)
                .setCancelable(false)
                .setPositiveButton("退出", (d, w) -> finish())
                .show();
    }

    void save(String url, String disp, String mime) {
        String name = null;
        if (disp != null) {
            String[] parts = disp.split(";");
            for (String p : parts) {
                p = p.trim();
                if (p.startsWith("filename*=UTF-8''")) {
                    try { name = URLDecoder.decode(p.substring(17), StandardCharsets.UTF_8.name()); } catch (Exception ignore) {}
                } else if (name == null && p.startsWith("filename=")) {
                    name = p.substring(9).replace("\"", "");
                }
            }
        }
        if (name == null || name.isEmpty()) { String[] seg = Uri.parse(url).getPath().split("/"); name = seg[seg.length - 1]; }
        name = name.replace("/", "_").replace("\", "_");
        try {
            DownloadManager dm = (DownloadManager) getSystemService(DOWNLOAD_SERVICE);
            DownloadManager.Request r = new DownloadManager.Request(Uri.parse(url));
            r.setDestinationInExternalPublicDir(Environment.DIRECTORY_DOWNLOADS, "能耗分析/" + name);
            r.setNotificationVisibility(DownloadManager.Request.VISIBILITY_VISIBLE_NOTIFY_COMPLETED);
            r.setTitle("能耗分析导出");
            r.setDescription(name);
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) r.setRequiresCharging(false);
            dm.enqueue(r);
            Toast.makeText(this, "已保存到「下载/能耗分析」文件夹", Toast.LENGTH_LONG).show();
        } catch (Exception e) {
            Toast.makeText(this, "保存失败：" + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    @Override
    protected void onActivityResult(int req, int res, Intent data) {
        super.onActivityResult(req, res, data);
        if (req == REQ_FILE) {
            ValueCallback<Uri[]> cb = fileCb;
            fileCb = null;
            if (cb == null) return;
            Uri u = (data != null && res == RESULT_OK) ? data.getData() : null;
            cb.onReceiveValue(u == null ? null : new Uri[]{u});
        }
    }

    @Override
    public void onBackPressed() {
        if (web != null && web.canGoBack() && !web.getUrl().endsWith("/")) {
            web.goBack();
        } else {
            super.onBackPressed();
        }
    }
}
