package mt.com.enemalta.paperless;

import android.net.Uri;
import android.os.Bundle;
import android.webkit.CookieManager;
import android.webkit.WebResourceRequest;
import android.webkit.WebView;

import com.getcapacitor.BridgeActivity;
import com.getcapacitor.BridgeWebViewClient;

public class MainActivity extends BridgeActivity {
  @Override
  public void onCreate(Bundle savedInstanceState) {
    super.onCreate(savedInstanceState);
    WebView webView = getBridge().getWebView();
    CookieManager cookieManager = CookieManager.getInstance();
    cookieManager.setAcceptCookie(true);
    cookieManager.setAcceptThirdPartyCookies(webView, true);

    webView.setWebViewClient(new BridgeWebViewClient(getBridge()) {
      private boolean loadingAppReturn;

      @Override
      public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
        Uri url = request.getUrl();
        String scheme = url.getScheme();
        if (!"https".equals(scheme) && !"http".equals(scheme)) {
          return super.shouldOverrideUrlLoading(view, request);
        }

        // A redirect to https://localhost is the app, not a website. If the
        // WebView follows it as a network request, Android reports
        // ERR_CONNECTION_REFUSED. Reload it through the local app server.
        if ("localhost".equalsIgnoreCase(url.getHost())) {
          if (loadingAppReturn) {
            loadingAppReturn = false;
            return false;
          }
          loadingAppReturn = true;
          String path = url.getPath();
          String target = (path == null || path.isEmpty())
            ? url.buildUpon().path("/").build().toString()
            : url.toString();
          view.loadUrl(target);
          return true;
        }

        // Keep Microsoft sign-in inside the app. Chrome cannot return here.
        return false;
      }
    });
  }
}
