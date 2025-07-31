// src/services/httpInterceptor.ts
import axios, { AxiosRequestConfig, AxiosResponse, AxiosError, InternalAxiosRequestConfig } from 'axios';
import { authService } from './authService';

interface AuthAxiosRequestConfig extends InternalAxiosRequestConfig {
    _retry?: boolean;
}

// 创建axios实例
export const createAuthenticatedAxiosInstance = (baseURL: string) => {
    const axiosInstance = axios.create({
        baseURL,
        timeout: 30000,
        headers: {
            'Content-Type': 'application/json'
        }
    });

    // 请求拦截器 - 自动添加认证token
    axiosInstance.interceptors.request.use(
        async (config: InternalAxiosRequestConfig): Promise<InternalAxiosRequestConfig> => {
            try {
                // 检查并刷新token（如果需要）
                const token = await authService.refreshTokenIfNeeded();
                
                if (token) {
                    // 确保headers对象存在
                    if (!config.headers) {
                        config.headers = {} as any;
                    }
                    config.headers.Authorization = `Bearer ${token}`;
                    
                    // 添加时间戳用于调试
                    console.log(`🔐 Request with token: ${config.method?.toUpperCase()} ${config.url}`);
                } else {
                    console.warn('⚠️  No token available for request:', config.url);
                }
                
                return config;
            } catch (error) {
                console.error('❌ Request interceptor error:', error);
                return Promise.reject(error);
            }
        },
        (error: AxiosError) => {
            console.error('❌ Request configuration error:', error);
            return Promise.reject(error);
        }
    );

    // 响应拦截器 - 处理认证错误和自动重试
    axiosInstance.interceptors.response.use(
        (response: AxiosResponse) => {
            // 请求成功，直接返回响应
            return response;
        },
        async (error: AxiosError) => {
            const originalRequest = error.config as AuthAxiosRequestConfig;
            
            // 处理401未授权错误
            if (error.response?.status === 401 && originalRequest && !originalRequest._retry) {
                originalRequest._retry = true;
                
                console.log('🔄 Token expired, attempting to refresh...');
                
                try {
                    // 尝试获取新token
                    const newToken = await authService.getTokenSilently();
                    
                    if (newToken) {
                        // 更新请求头
                        if (!originalRequest.headers) {
                            originalRequest.headers = {} as any;
                        }
                        originalRequest.headers.Authorization = `Bearer ${newToken}`;
                        
                        console.log('✅ Token refreshed, retrying request');
                        
                        // 重试原始请求
                        return axiosInstance(originalRequest);
                    } else {
                        // 无法获取新token，尝试交互式登录
                        console.log('🔓 Silent refresh failed, attempting interactive login...');
                        
                        const interactiveToken = await authService.loginInteractive();
                        
                        if (interactiveToken) {
                            if (!originalRequest.headers) {
                                originalRequest.headers = {} as any;
                            }
                            originalRequest.headers.Authorization = `Bearer ${interactiveToken}`;
                            return axiosInstance(originalRequest);
                        } else {
                            throw new Error('Authentication failed');
                        }
                    }
                } catch (refreshError) {
                    console.error('❌ Token refresh failed:', refreshError);
                    
                    // 认证彻底失败，清理状态并提示用户重新登录
                    await authService.logout();
                    
                    // 派发认证失败事件
                    window.dispatchEvent(new CustomEvent('authenticationFailed', {
                        detail: { error: 'Session expired. Please log in again.' }
                    }));
                    
                    return Promise.reject(new Error('Authentication failed. Please log in again.'));
                }
            }
            
            // 处理403禁止访问错误
            if (error.response?.status === 403) {
                console.error('❌ Access forbidden:', error.response.data);
                
                window.dispatchEvent(new CustomEvent('accessForbidden', {
                    detail: { error: 'Access denied. You do not have permission to access this resource.' }
                }));
            }
            
            // 处理网络错误
            if (!error.response) {
                console.error('❌ Network error:', error.message);
                
                window.dispatchEvent(new CustomEvent('networkError', {
                    detail: { error: 'Network connection failed. Please check your internet connection.' }
                }));
            }
            
            return Promise.reject(error);
        }
    );

    return axiosInstance;
};

// 为不同的API创建实例
export const apiClient = createAuthenticatedAxiosInstance(
    (() => {
        const hostname = window.location.hostname;
        
        if (hostname === "beone-d.beigenecorp.net" || hostname.includes("beigenecorp.net")) {
            return "https://beone-d.beigenecorp.net/api/aimta";
        }
        
        if (hostname === "localhost" || hostname === "127.0.0.1") {
            return "https://localhost:8000";
        }
        
        if (hostname === "10.8.63.207") {
            return "https://10.8.63.207:8000";
        }
        
        return "https://beone-d.beigenecorp.net/api/aimta";
    })()
);

// Token监控和自动刷新
class TokenMonitor {
    private refreshInterval: number | null = null;
    private readonly CHECK_INTERVAL = 5 * 60 * 1000; // 5分钟检查一次

    start() {
        if (this.refreshInterval) {
            return; // 已经在运行
        }

        this.refreshInterval = window.setInterval(async () => {
            try {
                const isAuthenticated = authService.isAuthenticated();
                if (isAuthenticated) {
                    await authService.refreshTokenIfNeeded();
                }
            } catch (error) {
                console.error('Token monitoring error:', error);
            }
        }, this.CHECK_INTERVAL);

        console.log('🔄 Token monitoring started');
    }

    stop() {
        if (this.refreshInterval) {
            window.clearInterval(this.refreshInterval);
            this.refreshInterval = null;
            console.log('⏹️  Token monitoring stopped');
        }
    }
}

export const tokenMonitor = new TokenMonitor();

// 在应用启动时开始监控
export const startTokenMonitoring = () => {
    tokenMonitor.start();
    
    // 在窗口获得焦点时检查token
    window.addEventListener('focus', async () => {
        try {
            const isAuthenticated = authService.isAuthenticated();
            if (isAuthenticated) {
                await authService.refreshTokenIfNeeded();
            }
        } catch (error) {
            console.error('Focus token check error:', error);
        }
    });
    
    // 在页面即将卸载时停止监控
    window.addEventListener('beforeunload', () => {
        tokenMonitor.stop();
    });
};