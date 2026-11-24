import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'vn.lite.workspace',
  appName: 'LITE Workspace',
  webDir: 'dist',
  android: {
    backgroundColor: '#FAF8F5',
  },
  plugins: {
    SplashScreen: {
      launchShowDuration: 900,
      launchAutoHide: true,
      backgroundColor: '#FAF8F5',
      androidScaleType: 'CENTER_CROP',
      showSpinner: false,
    },
    StatusBar: {
      style: 'DARK',
      backgroundColor: '#FAF8F5',
    },
  },
};

export default config;
