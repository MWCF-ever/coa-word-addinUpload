// src/taskpane/index.tsx - Updated with Auth Provider
import "../polyfills";
import React from "react";
import ReactDOM from "react-dom";
import { App } from "./components/App";
import { AuthProvider } from "../contexts/AuthContext";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { startTokenMonitoring } from "../services/httpInterceptor";

declare global {
  interface Window {
    __webpack_dev_server_client__?: boolean;
  }
}

/* global document, Office */
if (window.location.hostname.includes("officeapps.live.com")) {
  window.__webpack_dev_server_client__ = false;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log('🚀 Office.js ready, initializing AIMTA app...');
    console.log(`📋 Host: ${info.host}, Platform: ${info.platform}`);
    
    // 启动token监控
    startTokenMonitoring();
    
    ReactDOM.render(
      <React.StrictMode>
        <FluentProvider theme={webLightTheme}>
          <AuthProvider>
            <App />
          </AuthProvider>
        </FluentProvider>
      </React.StrictMode>,
      document.getElementById("container")
    );
    
    console.log('✅ AIMTA app initialized successfully');
  } else {
    console.error('❌ AIMTA app is only supported in Microsoft Word');
    
    // 显示错误信息
    ReactDOM.render(
      <React.StrictMode>
        <FluentProvider theme={webLightTheme}>
          <div style={{ 
            padding: '20px', 
            textAlign: 'center',
            color: '#d13438'
          }}>
            <h2>Unsupported Host Application</h2>
            <p>This add-in is designed specifically for Microsoft Word.</p>
            <p>Please open this add-in in Microsoft Word to continue.</p>
          </div>
        </FluentProvider>
      </React.StrictMode>,
      document.getElementById("container")
    );
  }
});