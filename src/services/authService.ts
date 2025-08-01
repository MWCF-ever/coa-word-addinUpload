// src/services/authService.ts - 修复版本
import { PublicClientApplication, Configuration, SilentRequest, PopupRequest, AuthenticationResult, AccountInfo } from "@azure/msal-browser";

interface AuthConfig {
    clientId: string;
    authority: string;
    redirectUri: string;
}

interface UserInfo {
    id: string;
    name: string;
    email: string;
    role?: string;
}

export class AuthService {
    private msalInstance: PublicClientApplication;
    private config: AuthConfig;
    private currentUser: UserInfo | null = null;

    constructor() {
        // 根据环境配置
        this.config = this.getAuthConfig();
        
        const msalConfig: Configuration = {
            auth: {
                clientId: this.config.clientId,
                authority: this.config.authority,
                redirectUri: this.config.redirectUri,
                navigateToLoginRequestUrl: false
            },
            cache: {
                cacheLocation: "sessionStorage", // 使用sessionStorage保持安全性
                storeAuthStateInCookie: true // 为IE11兼容性
            },
            system: {
                loggerOptions: {
                    loggerCallback: (level, message, containsPii) => {
                        if (containsPii) return;
                        console.log(`[MSAL] ${level}: ${message}`);
                    }
                }
            }
        };

        this.msalInstance = new PublicClientApplication(msalConfig);
        this.initialize();
    }

    private getAuthConfig(): AuthConfig {
        const hostname = window.location.hostname;
        
        // 根据环境确定配置
        const isDev = hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '10.8.63.207';
        
        return {
            clientId: "244a9262-04ff-4f5b-8958-2eeb0cedb928",
            authority: "https://login.microsoftonline.com/7dbc552d-50d7-4396-aeb9-04d0d393261b",
            // 修复重定向URI - 使用正确的SSO登录页面
            redirectUri: isDev 
                ? "https://localhost:3000/user/login"
                : "https://beone-d.beigenecorp.net/user/login"
        };
    }

    private async initialize(): Promise<void> {
        try {
            await this.msalInstance.initialize();
            
            // 检查URL参数中的token和action（SSO回调处理）
            this.handleUrlParameters();
            
            // 然后处理MSAL重定向
            await this.handleRedirectPromise();
            
        } catch (error) {
            console.error('MSAL initialization failed:', error);
        }
    }

    private async handleRedirectPromise(): Promise<void> {
        try {
            const response = await this.msalInstance.handleRedirectPromise();
            if (response && response.account) {
                this.msalInstance.setActiveAccount(response.account);
                await this.setCurrentUser(response);
            }
        } catch (error) {
            console.error('Redirect promise handling failed:', error);
        }
    }

    private handleUrlParameters(): void {
        const urlParams = new URLSearchParams(window.location.search);
        const token = urlParams.get('token');
        const refreshToken = urlParams.get('refreshToken');
        const action = urlParams.get('action');
        const redirect = urlParams.get('redirect');

        console.log('🔍 检查URL参数:', {
            action,
            hasToken: !!token,
            hasRefreshToken: !!refreshToken,
            redirect,
            fullUrl: window.location.href
        });

        if (action === 'login' && token) {
            console.log('✅ 检测到SSO回调，处理token...');
            
            // 存储token到sessionStorage
            sessionStorage.setItem('token', token);
            if (refreshToken) {
                sessionStorage.setItem('refreshToken', refreshToken);
            }
            
            // 获取用户信息
            this.fetchUserInfoFromToken(token);
            
            // 清理URL参数
            this.cleanUrlParameters();
            
            console.log('🎉 SSO登录处理完成');
            
            // 如果有redirect参数，可以用于后续导航
            if (redirect) {
                console.log('SSO登录成功，原始redirect目标:', redirect);
            }
        } else {
            console.log('🔍 未检测到SSO回调参数');
            // 检查是否有缓存的token
            const cachedToken = sessionStorage.getItem('token');
            if (cachedToken && !this.isTokenExpired(cachedToken)) {
                console.log('📋 发现缓存的token，尝试恢复用户状态');
                this.fetchUserInfoFromToken(cachedToken);
            } else if (cachedToken) {
                console.log('⚠️ 发现过期的token，清理缓存');
                sessionStorage.removeItem('token');
                sessionStorage.removeItem('refreshToken');
            }
        }
    }

    private async fetchUserInfoFromToken(token: string): Promise<void> {
        try {
            console.log('🔍 解析token获取用户信息...');
            
            // 解析JWT token获取用户信息
            const payload = this.parseJwtToken(token);
            
            console.log('📋 Token payload:', payload);
            
            // 更详细的用户信息提取
            this.currentUser = {
                id: payload.sub || payload.oid || payload.objectId,
                name: payload.name || payload.given_name + ' ' + payload.family_name || payload.preferred_username,
                email: payload.email || payload.preferred_username || payload.upn,
                role: payload.roles?.[0] || payload.role
            };

            // 验证用户信息的完整性
            if (!this.currentUser.email || this.currentUser.email === 'undefined') {
                console.warn('⚠️ 用户邮箱信息缺失，尝试其他字段');
                this.currentUser.email = payload.upn || payload.unique_name || '未知邮箱';
            }

            console.log('👤 用户信息设置完成:', this.currentUser);

            // 验证token有效性（调用后端验证）
            await this.validateTokenWithBackend(token);

            // 触发用户状态更新事件
            this.notifyUserStateChange();
        } catch (error) {
            console.error('❌ 解析token获取用户信息失败:', error);
            // 清理无效token
            sessionStorage.removeItem('token');
            sessionStorage.removeItem('refreshToken');
            this.currentUser = null;
        }
    }

    private async validateTokenWithBackend(token: string): Promise<void> {
        try {
            console.log('🔐 验证token有效性...');
            
            const response = await fetch(`${this.getApiBaseUrl()}/auth/status`, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                }
            });

            if (response.ok) {
                const data = await response.json();
                console.log('✅ Token验证成功:', data);
                
                // 如果后端返回了更完整的用户信息，更新当前用户
                if (data.user && data.authenticated) {
                    this.currentUser = {
                        ...this.currentUser,
                        ...data.user
                    };
                }
            } else {
                console.error('❌ Token验证失败:', response.status, response.statusText);
                throw new Error(`Token validation failed: ${response.status}`);
            }
        } catch (error) {
            console.error('❌ 后端token验证失败:', error);
            throw error;
        }
    }

    private getApiBaseUrl(): string {
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
    }

    private parseJwtToken(token: string): any {
        try {
            const base64Url = token.split('.')[1];
            const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
            const jsonPayload = decodeURIComponent(atob(base64).split('').map(c => 
                '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2)
            ).join(''));
            
            return JSON.parse(jsonPayload);
        } catch (error) {
            console.error('Failed to parse JWT token:', error);
            throw error;
        }
    }

    private cleanUrlParameters(): void {
        const url = new URL(window.location.href);
        url.searchParams.delete('token');
        url.searchParams.delete('refreshToken');
        url.searchParams.delete('action');
        url.searchParams.delete('redirect');
        window.history.replaceState({}, document.title, url.toString());
    }

    private notifyUserStateChange(): void {
        console.log('📢 触发用户状态更新事件:', this.currentUser);
        // 派发自定义事件通知应用状态更新
        window.dispatchEvent(new CustomEvent('userStateChanged', { 
            detail: { user: this.currentUser } 
        }));
    }

    // 公共方法：静默获取token
    public async getTokenSilently(scopes: string[] = ['User.Read']): Promise<string | null> {
        try {
            // 首先检查sessionStorage中的token
            const cachedToken = sessionStorage.getItem('token');
            if (cachedToken && !this.isTokenExpired(cachedToken)) {
                console.log('🔄 使用缓存的token');
                return cachedToken;
            }

            console.log('🔄 缓存token无效，尝试刷新...');

            // 尝试MSAL静默获取
            const account = this.msalInstance.getActiveAccount() || this.msalInstance.getAllAccounts()[0];
            if (!account) {
                console.log('❌ 没有活跃账户，无法静默获取token');
                return null;
            }

            const silentRequest: SilentRequest = {
                scopes,
                account
            };

            const response = await this.msalInstance.acquireTokenSilent(silentRequest);
            
            // 更新sessionStorage
            sessionStorage.setItem('token', response.accessToken);
            if (response.account) {
                await this.setCurrentUser(response);
            }

            console.log('✅ 静默token获取成功');
            return response.accessToken;
        } catch (error) {
            console.error('❌ Silent token acquisition failed:', error);
            return null;
        }
    }

    // 公共方法：交互式登录 - 跳转到SSO页面
    public async loginInteractive(scopes: string[] = ['User.Read']): Promise<string | null> {
        try {
            // 构建SSO登录URL
            const currentUrl = window.location.href;
            const redirectUrl = encodeURIComponent('/aimta/taskpane.html');
            const ssoLoginUrl = `https://beone-d.beigenecorp.net/user/login?redirect=${redirectUrl}`;
            
            console.log('跳转到SSO登录页面:', ssoLoginUrl);
            
            // 跳转到SSO登录页面
            window.location.href = ssoLoginUrl;
            
            // 注意：这里不会返回，因为页面会跳转
            return null;
        } catch (error) {
            console.error('SSO login redirect failed:', error);
            
            // 如果SSO跳转失败，fallback到MSAL popup登录
            try {
                const popupRequest: PopupRequest = {
                    scopes,
                    prompt: 'select_account'
                };

                const response = await this.msalInstance.loginPopup(popupRequest);
                
                if (response.account) {
                    this.msalInstance.setActiveAccount(response.account);
                    await this.setCurrentUser(response);
                    
                    // 存储token
                    sessionStorage.setItem('token', response.accessToken);
                    
                    return response.accessToken;
                }
                
                return null;
            } catch (fallbackError) {
                console.error('Fallback MSAL login also failed:', fallbackError);
                throw fallbackError;
            }
        }
    }

    // 公共方法：登出
    public async logout(): Promise<void> {
        try {
            // 清理sessionStorage
            sessionStorage.removeItem('token');
            sessionStorage.removeItem('refreshToken');
            
            // 清理用户状态
            this.currentUser = null;
            
            // MSAL登出
            const account = this.msalInstance.getActiveAccount();
            if (account) {
                await this.msalInstance.logoutPopup({
                    account,
                    postLogoutRedirectUri: this.config.redirectUri
                });
            }
            
            // 通知状态更新
            this.notifyUserStateChange();
        } catch (error) {
            console.error('Logout failed:', error);
        }
    }

    // 公共方法：获取当前用户
    public getCurrentUser(): UserInfo | null {
        return this.currentUser;
    }

    // 公共方法：检查是否已认证
    public isAuthenticated(): boolean {
        const token = sessionStorage.getItem('token');
        const hasActiveAccount = this.msalInstance.getAllAccounts().length > 0;
        const userExists = !!this.currentUser;
        
        const isAuth = (token && !this.isTokenExpired(token) && userExists) || hasActiveAccount;
        
        console.log('🔍 认证状态检查:', {
            hasToken: !!token,
            tokenValid: token ? !this.isTokenExpired(token) : false,
            hasActiveAccount,
            userExists,
            isAuthenticated: isAuth
        });
        
        return isAuth;
    }

    // Token续期检查
    public async refreshTokenIfNeeded(): Promise<string | null> {
        const token = sessionStorage.getItem('token');
        
        if (!token) {
            console.log('🔄 没有token，尝试静默获取');
            return await this.getTokenSilently();
        }

        // 检查token是否即将过期（5分钟内）
        if (this.isTokenExpiringSoon(token, 5 * 60 * 1000)) {
            console.log('🔄 Token即将过期，尝试刷新...');
            return await this.getTokenSilently();
        }

        return token;
    }

    private async setCurrentUser(response: AuthenticationResult): Promise<void> {
        if (response.account) {
            this.currentUser = {
                id: response.account.homeAccountId,
                name: response.account.name || response.account.username,
                email: response.account.username,
                role: (response.idTokenClaims as { roles?: string[] })?.roles?.[0]
            };
            this.notifyUserStateChange();
        }
    }

    private isTokenExpired(token: string): boolean {
        try {
            const payload = this.parseJwtToken(token);
            const currentTime = Math.floor(Date.now() / 1000);
            const isExpired = payload.exp < currentTime;
            
            if (isExpired) {
                console.log('⚠️ Token已过期');
            }
            
            return isExpired;
        } catch {
            console.log('⚠️ Token解析失败，视为过期');
            return true;
        }
    }

    private isTokenExpiringSoon(token: string, thresholdMs: number): boolean {
        try {
            const payload = this.parseJwtToken(token);
            const currentTime = Math.floor(Date.now() / 1000);
            const expirationTime = payload.exp;
            const timeUntilExpiry = (expirationTime - currentTime) * 1000;
            
            return timeUntilExpiry <= thresholdMs;
        } catch {
            return true;
        }
    }
}

// 单例模式
export const authService = new AuthService();