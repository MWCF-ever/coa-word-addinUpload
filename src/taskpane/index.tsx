// src/taskpane/index.tsx - 修复版本，包含调试工具加载
import "../polyfills";
import React from "react";
import ReactDOM from "react-dom";
import { App } from "./components/App";
import { AuthProvider } from "../contexts/AuthContext";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { startTokenMonitoring } from "../services/httpInterceptor";

// 🔧 修复：导入并初始化Office加载项调试工具
import "../utils/officeAddinHelpers";

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
    console.log(`🌐 Environment: ${window.location.hostname}`);
    console.log(`🏢 Office Environment: ${window.isOfficeAddinEnvironment?.() ? 'Yes' : 'No'}`);
    
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
    
    // 🔧 在开发模式下显示调试信息
    if (process.env.NODE_ENV === 'development') {
      setTimeout(() => {
        console.log('🔧 Office加载项调试工具已就绪');
        console.log('💡 可用的调试命令:');
        console.log('  - window.isOfficeAddinEnvironment() - 检查Office环境');
        console.log('  - window.testOfficeAddinConnection() - 测试连接');
        console.log('  - window.fullOfficeAddinTest() - 完整测试');
        console.log('  - window.officeAddinApiClient - API客户端');
        
        // 自动运行一次基本检查
        if (window.isOfficeAddinEnvironment?.()) {
          console.log('🏢 检测到Office环境，运行自动诊断...');
          window.testOfficeAddinConnection?.().then(result => {
            console.log('📊 自动诊断结果:', result);
          }).catch(error => {
            console.error('❌ 自动诊断失败:', error);
          });
        }
      }, 2000);
    }
    
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
            <p><small>Current host: {info.host}</small></p>
          </div>
        </FluentProvider>
      </React.StrictMode>,
      document.getElementById("container")
    );
  }
});