// src/components/AuthGuard.tsx
import React, { useEffect, useState } from 'react';
import { Stack, Spinner, SpinnerSize, PrimaryButton, Text, MessageBar, MessageBarType } from '@fluentui/react';
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

    useEffect(() => {
        // 监听认证失败事件
        const handleAuthFailed = (event: CustomEvent) => {
            setShowRetry(true);
        };

        const handleAccessForbidden = (event: CustomEvent) => {
            console.error('Access forbidden:', event.detail.error);
        };

        const handleNetworkError = (event: CustomEvent) => {
            console.error('Network error:', event.detail.error);
        };

        window.addEventListener('authenticationFailed', handleAuthFailed as EventListener);
        window.addEventListener('accessForbidden', handleAccessForbidden as EventListener);
        window.addEventListener('networkError', handleNetworkError as EventListener);

        return () => {
            window.removeEventListener('authenticationFailed', handleAuthFailed as EventListener);
            window.removeEventListener('accessForbidden', handleAccessForbidden as EventListener);
            window.removeEventListener('networkError', handleNetworkError as EventListener);
        };
    }, []);

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
                    <PrimaryButton
                        text="Try Again"
                        iconProps={{ iconName: 'Refresh' }}
                        onClick={login}
                    />
                </Stack>
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
                    {user.email}
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