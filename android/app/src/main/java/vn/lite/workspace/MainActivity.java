package vn.lite.workspace;

import android.content.ClipData;
import android.content.Intent;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        normalizeIncomingFileIntent(getIntent());
        super.onCreate(savedInstanceState);
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
