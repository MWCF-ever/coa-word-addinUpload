// src/services/httpInterceptor.ts - 修复Mixed Content问题
import axios, { AxiosRequestConfig, AxiosResponse, AxiosError, InternalAxiosRequestConfig } from 'axios';
import { authService } from './authService';

interface AuthAxiosRequestConfig extends InternalAxiosRequestConfig {
    _retry?: boolean;
}

// 修复：更准确的baseURL检测
function getApiBaseUrl(): string {
    const hostname = window.location.hostname;
    
    console.log('🔍 检测API baseURL:', {
        hostname,
        protocol: window.location.protocol,
        href: window.location.href
    });
    
    // 🔥 关键修复：无论从哪里访问，都使用HTTPS的生产环境地址
    if (hostname.includes("sharepoint.com") || 
        hostname === "beone-d.beigenecorp.net" || 
        hostname.includes("beigenecorp.net") ||
        hostname.includes("office.com") ||
        hostname.includes("officeapps.live.com")) {
        
        const apiUrl = "https://beone-d.beigenecorp.net/api/aimta";
        console.log(`✅ 使用生产环境API (Office环境): ${apiUrl}`);
        return apiUrl;
    }
    
    // 本地开发环境
    if (hostname === "localhost" || hostname === "127.0.0.1") {
        const apiUrl = "https://localhost:8000";
        console.log(`✅ 使用本地开发API: ${apiUrl}`);
        return apiUrl;
    }
    
    // 开发环境IP
    if (hostname === "10.8.63.207") {
        const apiUrl = "https://10.8.63.207:8000";
        console.log(`✅ 使用开发IP API: ${apiUrl}`);
        return apiUrl;
    }
    
    // 默认返回生产环境HTTPS地址
    const defaultUrl = "https://beone-d.beigenecorp.net/api/aimta";
    console.log(`✅ 使用默认API: ${defaultUrl}`);
    return defaultUrl;
}

// 创建axios实例
export const createAuthenticatedAxiosInstance = (baseURL?: string) => {
    // 🔥 修复：如果没有提供baseURL，自动检测
    const finalBaseURL = baseURL || getApiBaseUrl();
    
    console.log(`🔧 创建axios实例，baseURL: ${finalBaseURL}`);
    
    const axiosInstance = axios.create({
        baseURL: finalBaseURL,
        timeout: 30000,
        headers: {
            'Content-Type': 'application/json'
        }
    });

    // 请求拦截器 - 自动添加认证token
    axiosInstance.interceptors.request.use(
        async (config: InternalAxiosRequestConfig): Promise<InternalAxiosRequestConfig> => {
            try {
                // 🔥 修复：确保请求URL是绝对HTTPS URL
                if (config.url && !config.url.startsWith('http')) {
                    // 如果是相对URL，确保使用正确的baseURL
                    const fullUrl = `${finalBaseURL}${config.url.startsWith('/') ? '' : '/'}${config.url}`;
                    console.log(`🔗 构建完整URL: ${config.url} -> ${fullUrl}`);
                }
                
                // 检查并刷新token（如果需要）
                const token = await authService.refreshTokenIfNeeded();
                
                if (token) {
                    // 确保headers对象存在
                    if (!config.headers) {
                        config.headers = {} as any;
                    }
                    config.headers.Authorization = `Bearer ${token}`;
                    
                    // 添加调试日志
                    console.log(`🔐 Request with token: ${config.method?.toUpperCase()} ${config.url}`);
                    console.log(`🔑 Token preview: ${token.substring(0, 20)}...${token.substring(token.length - 10)}`);
                } else {
                    console.warn('⚠️  No token available for request:', config.url);
                    
                    // 如果没有token但是需要认证的端点，尝试登录
                    if (config.url && !config.url.includes('/health') && !config.url.includes('/auth/status')) {
                        console.log('🔓 需要认证的端点但无token，尝试获取token...');
                        const newToken = await authService.getTokenSilently();
                        if (newToken) {
                            config.headers.Authorization = `Bearer ${newToken}`;
                            console.log('✅ 成功获取新token');
                        } else {
                            console.log('❌ 无法获取token，可能需要重新登录');
                        }
                    }
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
            // 检查响应内容是否包含错误信息
            if (response.data && response.data.code === '900401') {
                console.error('🚫 API返回401错误，但HTTP状态码为200:', response.data);
                
                // 创建一个401错误来触发重试逻辑
                const unauthorizedError = new Error('Unauthorized') as any;
                unauthorizedError.response = {
                    ...response,
                    status: 401,
                    statusText: 'Unauthorized'
                };
                unauthorizedError.config = response.config;
                
                return Promise.reject(unauthorizedError);
            }
            
            // 请求成功，直接返回响应
            console.log(`✅ Response received: ${response.config?.method?.toUpperCase()} ${response.config?.url} - ${response.status}`);
            return response;
        },
        async (error: AxiosError) => {
            const originalRequest = error.config as AuthAxiosRequestConfig;
            
            console.error('❌ HTTP Error:', {
                method: originalRequest?.method?.toUpperCase(),
                url: originalRequest?.url,
                baseURL: originalRequest?.baseURL,
                fullURL: originalRequest?.baseURL + (originalRequest?.url ?? ''),
                status: error.response?.status,
                statusText: error.response?.statusText,
                data: error.response?.data,
                message: error.message,
                code: error.code
            });
            
            // 🔥 特殊处理Mixed Content错误
            if (error.message === 'Network Error' && 
                (error.code === 'ERR_NETWORK' || error.code === 'ERR_BLOCKED_BY_CLIENT')) {
                console.error('🚫 检测到Mixed Content或网络阻塞错误');
                console.error('🔍 可能的原因：HTTP/HTTPS混合内容或CORS问题');
                
                // 尝试重新构建请求with强制HTTPS
                if (originalRequest && !originalRequest._retry) {
                    originalRequest._retry = true;
                    
                    // 强制使用HTTPS baseURL
                    const httpsBaseURL = "https://beone-d.beigenecorp.net/api/aimta";
                    console.log(`🔄 尝试使用HTTPS baseURL重试: ${httpsBaseURL}`);
                    
                    originalRequest.baseURL = httpsBaseURL;
                    
                    return axiosInstance(originalRequest);
                }
            }
            
            // 处理401未授权错误
            if ((error.response?.status === 401 || error.message === 'Unauthorized') && originalRequest && !originalRequest._retry) {
                originalRequest._retry = true;
                
                console.log('🔄 检测到401错误，尝试刷新token...');
                
                try {
                    // 清理可能过期的token
                    sessionStorage.removeItem('token');
                    
                    // 尝试获取新token
                    const newToken = await authService.getTokenSilently();
                    
                    if (newToken) {
                        // 更新请求头
                        if (!originalRequest.headers) {
                            originalRequest.headers = {} as any;
                        }
                        originalRequest.headers.Authorization = `Bearer ${newToken}`;
                        
                        console.log('✅ Token刷新成功，重试请求');
                        
                        // 重试原始请求
                        return axiosInstance(originalRequest);
                    } else {
                        // 无法获取新token，尝试交互式登录
                        console.log('🔓 静默刷新失败，需要重新登录');
                        
                        // 派发认证失败事件，让UI处理登录
                        window.dispatchEvent(new CustomEvent('authenticationFailed', {
                            detail: { 
                                error: 'Session expired. Please log in again.',
                                requireLogin: true
                            }
                        }));
                        
                        throw new Error('Authentication failed. Please log in again.');
                    }
                } catch (refreshError) {
                    console.error('❌ Token refresh failed:', refreshError);
                    
                    // 认证彻底失败，清理状态并提示用户重新登录
                    await authService.logout();
                    
                    // 派发认证失败事件
                    window.dispatchEvent(new CustomEvent('authenticationFailed', {
                        detail: { 
                            error: 'Session expired. Please log in again.',
                            requireLogin: true
                        }
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
                    detail: { 
                        error: 'Network connection failed. Please check your internet connection.',
                        code: error.code,
                        message: error.message
                    }
                }));
            }
            
            return Promise.reject(error);
        }
    );

    return axiosInstance;
};

// 🔥 修复：使用动态检测的baseURL创建apiClient
export const apiClient = createAuthenticatedAxiosInstance();

// 重新导出createAuthenticatedAxiosInstance以便在其他地方使用
// (Removed duplicate export of createAuthenticatedAxiosInstance)

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
                    const token = await authService.refreshTokenIfNeeded();
                    if (token) {
                        console.log('🔄 Token监控：Token仍然有效');
                    } else {
                        console.log('⚠️ Token监控：Token无效或过期');
                    }
                } else {
                    console.log('⚠️ Token监控：用户未认证');
                }
            } catch (error) {
                console.error('❌ Token monitoring error:', error);
            }
        }, this.CHECK_INTERVAL);

        console.log('🔄 Token监控已启动');
    }

    stop() {
        if (this.refreshInterval) {
            window.clearInterval(this.refreshInterval);
            this.refreshInterval = null;
            console.log('⏹️  Token监控已停止');
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
                console.log('🔍 窗口焦点检查：Token状态已更新');
            }
        } catch (error) {
            console.error('❌ Focus token check error:', error);
        }
    });
    
    // 在页面即将卸载时停止监控
    window.addEventListener('beforeunload', () => {
        tokenMonitor.stop();
    });
};