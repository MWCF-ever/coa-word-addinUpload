import "../polyfills";
import React from "react";
import ReactDOM from "react-dom";
import { App } from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

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
    ReactDOM.render(
      <React.StrictMode>
        <FluentProvider theme={webLightTheme}>
          <App />
        </FluentProvider>
      </React.StrictMode>,
      document.getElementById("container")
    );
  }
});