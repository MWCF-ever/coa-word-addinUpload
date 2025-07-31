// src/types/global.d.ts
import { AxiosRequestHeaders } from 'axios';

declare global {
  interface Window {
    __webpack_dev_server_client__?: boolean;
  }
}

// 扩展 AxiosRequestConfig 以包含 _retry 属性
declare module 'axios' {
  export interface InternalAxiosRequestConfig {
    _retry?: boolean;
  }
}

export {};