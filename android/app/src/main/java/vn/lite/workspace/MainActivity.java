package vn.lite.workspace;

import android.content.ClipData;
import android.content.Intent;
import android.graphics.Color;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.view.Window;
import android.webkit.JavascriptInterface;
import androidx.core.view.WindowCompat;
import androidx.core.view.WindowInsetsControllerCompat;
import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        normalizeIncomingFileIntent(getIntent());
        super.onCreate(savedInstanceState);
        applyStatusBarAppearance("#FAF8F5", false);
        if (this.bridge != null && this.bridge.getWebView() != null) {
            this.bridge.getWebView().addJavascriptInterface(
                new LITEAndroidSystemBridge(),
                "LITEAndroidSystem"
            );
        }
    }

    private void applyStatusBarAppearance(String colorHex, boolean isDarkBackground) {
        try {
            Window window = getWindow();
            if (window == null) {
                return;
            }
            window.setStatusBarColor(Color.parseColor(colorHex));
            WindowInsetsControllerCompat controller =
                WindowCompat.getInsetsController(window, window.getDecorView());
            if (controller != null) {
                controller.setAppearanceLightStatusBars(!isDarkBackground);
            }
        } catch (IllegalArgumentException ignored) {
            // Ignore malformed color string
        }
    }

    private class LITEAndroidSystemBridge {
        @JavascriptInterface
        public void setStatusBarTheme(String colorHex, boolean isDarkBackground) {
            runOnUiThread(() -> applyStatusBarAppearance(colorHex, isDarkBackground));
        }
    }

    @Override
    protected void onNewIntent(Intent intent) {
        normalizeIncomingFileIntent(intent);
        if (intent != null) {
            setIntent(intent);
        }
        super.onNewIntent(intent);
    }

    /**
     * Normalizes ACTION_SEND file intents (e.g. shared from Zalo/Messenger via EXTRA_STREAM or ClipData)
     * into ACTION_VIEW with data URI + read permission preserved so Capacitor App plugin can surface
     * them through getLaunchUrl() and appUrlOpen.
     */
    private void normalizeIncomingFileIntent(Intent intent) {
        if (intent == null) {
            return;
        }

        if (Intent.ACTION_SEND.equals(intent.getAction())) {
            Uri uri = extractSharedUri(intent);
            if (uri != null) {
                String mimeType = intent.getType();
                intent.setAction(Intent.ACTION_VIEW);
                if (mimeType != null && !mimeType.isEmpty()) {
                    intent.setDataAndType(uri, mimeType);
                } else {
                    intent.setData(uri);
                }
                intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                setIntent(intent);
            }
        } else if (Intent.ACTION_VIEW.equals(intent.getAction()) && intent.getData() != null) {
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        }
    }

    private Uri extractSharedUri(Intent intent) {
        Uri uri;
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU) {
            uri = intent.getParcelableExtra(Intent.EXTRA_STREAM, Uri.class);
        } else {
            uri = getLegacyParcelableExtra(intent, Intent.EXTRA_STREAM);
        }

        if (uri == null) {
            ClipData clipData = intent.getClipData();
            if (clipData != null && clipData.getItemCount() > 0) {
                uri = clipData.getItemAt(0).getUri();
            }
        }

        if (uri == null) {
            uri = intent.getData();
        }

        return uri;
    }

    @SuppressWarnings("deprecation")
    private Uri getLegacyParcelableExtra(Intent intent, String key) {
        return intent.getParcelableExtra(key);
    }
}
