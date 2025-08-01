// src/components/AuthGuard.tsx - 增强调试版本
import React, { useEffect, useState } from 'react';
import { Stack, Spinner, SpinnerSize, PrimaryButton, Text, MessageBar, MessageBarType, DefaultButton } from '@fluentui/react';
import { useAuth } from '../contexts/AuthContext';

interface AuthGuardProps {
    children: React.ReactNode;
    fallback?: React.ReactNode;
    requireAuth?: boolean;
}

export const AuthGuard: React.FC<AuthGuardProps> = ({ 
    children, 
    fallback,
    requireAuth = true 
}) => {
    const { user, isAuthenticated, isLoading, error, login, clearError } = useAuth();
    const [showRetry, setShowRetry] = useState(false);
    const [debugInfo, setDebugInfo] = useState<any>(null);

    useEffect(() => {
        // 收集调试信息
        const collectDebugInfo = () => {
            const token = sessionStorage.getItem('token');
            const refreshToken = sessionStorage.getItem('refreshToken');
            
            setDebugInfo({
                hasToken: !!token,
                tokenPreview: token ? `${token.substring(0, 20)}...` : null,
                hasRefreshToken: !!refreshToken,
                userInfo: user,
                isAuthenticated,
                isLoading,
                error,
                url: window.location.href,
                timestamp: new Date().toISOString()
            });
        };

        collectDebugInfo();

        // 监听认证失败事件
        const handleAuthFailed = (event: CustomEvent) => {
            console.error('🚫 认证失败事件:', event.detail);
            setShowRetry(true);
            collectDebugInfo();
        };

        const handleAccessForbidden = (event: CustomEvent) => {
            console.error('🚫 访问被禁止:', event.detail.error);
            collectDebugInfo();
        };

        const handleNetworkError = (event: CustomEvent) => {
            console.error('🌐 网络错误:', event.detail.error);
            collectDebugInfo();
        };

        window.addEventListener('authenticationFailed', handleAuthFailed as EventListener);
        window.addEventListener('accessForbidden', handleAccessForbidden as EventListener);
        window.addEventListener('networkError', handleNetworkError as EventListener);

        // 定期更新调试信息
        const debugInterval = setInterval(collectDebugInfo, 10000); // 每10秒更新一次

        return () => {
            window.removeEventListener('authenticationFailed', handleAuthFailed as EventListener);
            window.removeEventListener('accessForbidden', handleAccessForbidden as EventListener);
            window.removeEventListener('networkError', handleNetworkError as EventListener);
            clearInterval(debugInterval);
        };
    }, [user, isAuthenticated, isLoading, error]);

    // 强制刷新认证状态
    const handleForceRefresh = async () => {
        console.log('🔄 强制刷新认证状态...');
        try {
            // 清理当前状态
            sessionStorage.removeItem('token');
            sessionStorage.removeItem('refreshToken');
            
            // 重新尝试登录
            await login();
        } catch (error) {
            console.error('❌ 强制刷新失败:', error);
        }
    };

    // 清理所有数据
    const handleClearAll = () => {
        console.log('🧹 清理所有认证数据...');
        sessionStorage.clear();
        localStorage.clear();
        window.location.reload();
    };

    // 如果不需要认证，直接渲染子组件
    if (!requireAuth) {
        return <>{children}</>;
    }

    // 加载中状态
    if (isLoading) {
        return (
            <Stack horizontalAlign="center" verticalAlign="center" 
                   tokens={{ padding: 40 }} styles={{ root: { minHeight: '200px' } }}>
                <Spinner size={SpinnerSize.large} label="Authenticating..." />
                <Text variant="medium" styles={{ root: { marginTop: 16 } }}>
                    Please wait while we verify your credentials...
                </Text>
                
                {/* 调试信息 */}
                {debugInfo && (
                    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 20, padding: 10, backgroundColor: '#f5f5f5', borderRadius: 4 } }}>
                        <Text variant="small" styles={{ root: { fontWeight: 600 } }}>调试信息:</Text>
                        <Text variant="small">Token存在: {debugInfo.hasToken ? '✅' : '❌'}</Text>
                        <Text variant="small">认证状态: {debugInfo.isAuthenticated ? '✅' : '❌'}</Text>
                        <Text variant="small">用户信息: {debugInfo.userInfo ? '✅' : '❌'}</Text>
                        {debugInfo.tokenPreview && (
                            <Text variant="small">Token预览: {debugInfo.tokenPreview}</Text>
                        )}
                        <Text variant="small">时间: {new Date(debugInfo.timestamp).toLocaleTimeString()}</Text>
                    </Stack>
                )}
            </Stack>
        );
    }

    // 认证错误状态
    if (error) {
        return (
            <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: 20 } }}>
                <MessageBar 
                    messageBarType={MessageBarType.error}
                    onDismiss={clearError}
                >
                    {error}
                </MessageBar>
                
                <Stack horizontalAlign="center" tokens={{ childrenGap: 15 }}>
                    <Text variant="large">Authentication Error</Text>
                    <Text variant="medium">
                        We encountered an issue while trying to authenticate you.
                    </Text>
                    
                    <Stack horizontal tokens={{ childrenGap: 10 }}>
                        <PrimaryButton
                            text="Try Again"
                            iconProps={{ iconName: 'Refresh' }}
                            onClick={login}
                        />
                        <DefaultButton
                            text="Force Refresh"
                            iconProps={{ iconName: 'Sync' }}
                            onClick={handleForceRefresh}
                        />
                        <DefaultButton
                            text="Clear All"
                            iconProps={{ iconName: 'Delete' }}
                            onClick={handleClearAll}
                        />
                    </Stack>
                </Stack>

                {/* 详细调试信息 */}
                {debugInfo && (
                    <Stack tokens={{ childrenGap: 5 }} styles={{ root: { padding: 15, backgroundColor: '#f8f8f8', borderRadius: 4, fontSize: '12px' } }}>
                        <Text variant="small" styles={{ root: { fontWeight: 600, color: '#d13438' } }}>详细调试信息:</Text>
                        <Text variant="small">URL: {debugInfo.url}</Text>
                        <Text variant="small">Has Token: {debugInfo.hasToken ? 'Yes' : 'No'}</Text>
                        <Text variant="small">Has Refresh Token: {debugInfo.hasRefreshToken ? 'Yes' : 'No'}</Text>
                        <Text variant="small">Is Authenticated: {debugInfo.isAuthenticated ? 'Yes' : 'No'}</Text>
                        <Text variant="small">Is Loading: {debugInfo.isLoading ? 'Yes' : 'No'}</Text>
                        <Text variant="small">Error: {debugInfo.error || 'None'}</Text>
                        <Text variant="small">User: {JSON.stringify(debugInfo.userInfo, null, 2)}</Text>
                        <Text variant="small">Timestamp: {debugInfo.timestamp}</Text>
                    </Stack>
                )}
            </Stack>
        );
    }

    // 未认证状态
    if (!isAuthenticated || !user) {
        if (fallback) {
            return <>{fallback}</>;
        }

        return (
            <Stack horizontalAlign="center" verticalAlign="center" 
                   tokens={{ childrenGap: 20 }} 
                   styles={{ root: { minHeight: '300px', padding: 20 } }}>
                
                <Stack horizontalAlign="center" tokens={{ childrenGap: 15 }}>
                    <div style={{ 
                        fontSize: '48px', 
                        color: '#0078d4',
                        marginBottom: '10px'
                    }}>
                        🔐
                    </div>
                    
                    <Text variant="xxLarge" styles={{ root: { fontWeight: 600 } }}>
                        AIMTA Document Processor
                    </Text>
                    
                    <Text variant="large" styles={{ root: { textAlign: 'center', maxWidth: '400px' } }}>
                        Please sign in with your corporate account to access the AIMTA batch analysis features.
                    </Text>
                    
                    {showRetry && (
                        <MessageBar messageBarType={MessageBarType.warning}>
                            Your session has expired. Please sign in again to continue.
                        </MessageBar>
                    )}
                    
                    <Stack horizontal tokens={{ childrenGap: 10 }}>
                        <PrimaryButton
                            text="Sign In"
                            iconProps={{ iconName: 'Signin' }}
                            onClick={login}
                            styles={{ 
                                root: { 
                                    minWidth: '140px',
                                    height: '40px',
                                    fontSize: '16px'
                                } 
                            }}
                        />
                        <DefaultButton
                            text="Force Refresh"
                            iconProps={{ iconName: 'Sync' }}
                            onClick={handleForceRefresh}
                        />
                    </Stack>
                </Stack>
                
                <Stack className="help-text" tokens={{ childrenGap: 5 }}>
                    <Text variant="small" styles={{ root: { textAlign: 'center' } }}>
                        • Use your BeiGene corporate credentials
                    </Text>
                    <Text variant="small" styles={{ root: { textAlign: 'center' } }}>
                        • Single Sign-On (SSO) supported
                    </Text>
                    <Text variant="small" styles={{ root: { textAlign: 'center' } }}>
                        • Secure authentication via Microsoft Entra ID
                    </Text>
                </Stack>

                {/* 调试信息面板 */}
                {debugInfo && (
                    <details style={{ width: '100%', maxWidth: '600px' }}>
                        <summary style={{ cursor: 'pointer', padding: '10px', backgroundColor: '#e1dfdd', borderRadius: '4px' }}>
                            🔧 调试信息 (点击展开)
                        </summary>
                        <Stack tokens={{ childrenGap: 5 }} styles={{ root: { padding: 15, backgroundColor: '#f8f8f8', borderRadius: 4, fontSize: '12px', marginTop: 10 } }}>
                            <Text variant="small">Current URL: {debugInfo.url}</Text>
                            <Text variant="small">Has Token: {debugInfo.hasToken ? 'Yes' : 'No'}</Text>
                            {debugInfo.tokenPreview && (
                                <Text variant="small">Token Preview: {debugInfo.tokenPreview}</Text>
                            )}
                            <Text variant="small">Has Refresh Token: {debugInfo.hasRefreshToken ? 'Yes' : 'No'}</Text>
                            <Text variant="small">Is Authenticated: {debugInfo.isAuthenticated ? 'Yes' : 'No'}</Text>
                            <Text variant="small">User Info: {debugInfo.userInfo ? JSON.stringify(debugInfo.userInfo) : 'None'}</Text>
                            <Text variant="small">Last Updated: {new Date(debugInfo.timestamp).toLocaleString()}</Text>
                            
                            <Stack horizontal tokens={{ childrenGap: 5 }} styles={{ root: { marginTop: 10 } }}>
                                <DefaultButton
                                    text="检查Token"
                                    onClick={() => {
                                        const token = sessionStorage.getItem('token');
                                        console.log('Current token:', token);
                                        alert(token ? `Token exists: ${token.substring(0, 50)}...` : 'No token found');
                                    }}
                                />
                                <DefaultButton
                                    text="检查URL参数"
                                    onClick={() => {
                                        const params = new URLSearchParams(window.location.search);
                                        const info = Array.from(params.entries()).map(([key, value]) => `${key}: ${value}`).join('\n');
                                        console.log('URL parameters:', info);
                                        alert(info || 'No URL parameters found');
                                    }}
                                />
                                <DefaultButton
                                    text="清理缓存"
                                    onClick={handleClearAll}
                                />
                            </Stack>
                        </Stack>
                    </details>
                )}
            </Stack>
        );
    }

    // 已认证，渲染子组件
    return <>{children}</>;
};

// 用户信息显示组件
export const UserInfo: React.FC = () => {
    const { user, logout } = useAuth();

    if (!user) {
        return null;
    }

    return (
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center"
               styles={{ 
                   root: { 
                       padding: '8px 16px', 
                       backgroundColor: '#f3f2f1',
                       borderRadius: '4px',
                       marginBottom: '16px'
                   } 
               }}>
            <Stack tokens={{ childrenGap: 2 }}>
                <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                    {user.name}
                </Text>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    {user.email || '邮箱未知'}
                </Text>
                <Text variant="small" styles={{ root: { color: '#605e5c', fontSize: '10px' } }}>
                    ID: {user.id ? user.id.substring(0, 8) + '...' : '未知'}
                </Text>
            </Stack>
            
            <PrimaryButton
                text="Sign Out"
                iconProps={{ iconName: 'SignOut' }}
                onClick={logout}
                styles={{ 
                    root: { 
                        minWidth: '80px',
                        height: '32px'
                    } 
                }}
            />
        </Stack>
    );
};