// src/contexts/AuthContext.tsx
import React, { createContext, useContext, useState, useEffect, ReactNode } from 'react';
import { authService } from '../services/authService';

interface UserInfo {
    id: string;
    name: string;
    email: string;
    role?: string;
}

interface AuthContextType {
    user: UserInfo | null;
    isAuthenticated: boolean;
    isLoading: boolean;
    error: string | null;
    login: () => Promise<void>;
    logout: () => Promise<void>;
    getToken: () => Promise<string | null>;
    clearError: () => void;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

interface AuthProviderProps {
    children: ReactNode;
}

export const AuthProvider: React.FC<AuthProviderProps> = ({ children }) => {
    const [user, setUser] = useState<UserInfo | null>(null);
    const [isAuthenticated, setIsAuthenticated] = useState(false);
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        initializeAuth();
        
        // 监听用户状态变化事件
        const handleUserStateChange = (event: CustomEvent) => {
            const userData = event.detail.user;
            console.log('📢 接收到用户状态变化事件:', userData);
            setUser(userData);
            setIsAuthenticated(!!userData);
            
            // 如果用户登录成功，清除错误状态
            if (userData) {
                setError(null);
            }
        };

        window.addEventListener('userStateChanged', handleUserStateChange as EventListener);
        
        return () => {
            window.removeEventListener('userStateChanged', handleUserStateChange as EventListener);
        };
    }, []);

    const initializeAuth = async () => {
        try {
            setIsLoading(true);
            setError(null);

            console.log('🔄 初始化认证状态...');

            // 检查是否已认证
            const authenticated = authService.isAuthenticated();
            const currentUser = authService.getCurrentUser();
            
            console.log('🔍 认证检查结果:', { authenticated, currentUser });
            
            if (authenticated && currentUser) {
                console.log('✅ 用户已认证:', currentUser);
                setUser(currentUser);
                setIsAuthenticated(true);
            } else {
                console.log('⚠️ 用户未认证，尝试静默获取token...');
                // 尝试静默获取token
                const token = await authService.getTokenSilently();
                if (token) {
                    const updatedUser = authService.getCurrentUser();
                    console.log('✅ 静默登录成功:', updatedUser);
                    setUser(updatedUser);
                    setIsAuthenticated(!!updatedUser);
                } else {
                    console.log('❌ 静默登录失败');
                    setUser(null);
                    setIsAuthenticated(false);
                }
            }
        } catch (err) {
            console.error('❌ 认证初始化失败:', err);
            setError('Authentication initialization failed');
            setUser(null);
            setIsAuthenticated(false);
        } finally {
            setIsLoading(false);
        }
    };

    const login = async () => {
        try {
            setIsLoading(true);
            setError(null);

            // 首先尝试静默登录
            let token = await authService.getTokenSilently();
            
            if (!token) {
                // 静默登录失败，使用交互式登录
                token = await authService.loginInteractive();
            }

            if (token) {
                const currentUser = authService.getCurrentUser();
                setUser(currentUser);
                setIsAuthenticated(!!currentUser);
            } else {
                throw new Error('Login failed');
            }
        } catch (err) {
            console.error('Login failed:', err);
            setError('Login failed. Please try again.');
            setIsAuthenticated(false);
            setUser(null);
        } finally {
            setIsLoading(false);
        }
    };

    const logout = async () => {
        try {
            setIsLoading(true);
            setError(null);

            await authService.logout();
            setUser(null);
            setIsAuthenticated(false);
        } catch (err) {
            console.error('Logout failed:', err);
            setError('Logout failed');
        } finally {
            setIsLoading(false);
        }
    };

    const getToken = async (): Promise<string | null> => {
        try {
            // 自动检查并刷新token
            return await authService.refreshTokenIfNeeded();
        } catch (err) {
            console.error('Token retrieval failed:', err);
            setError('Token retrieval failed');
            return null;
        }
    };

    const clearError = () => {
        setError(null);
    };

    const contextValue: AuthContextType = {
        user,
        isAuthenticated,
        isLoading,
        error,
        login,
        logout,
        getToken,
        clearError
    };

    return (
        <AuthContext.Provider value={contextValue}>
            {children}
        </AuthContext.Provider>
    );
};

export const useAuth = (): AuthContextType => {
    const context = useContext(AuthContext);
    if (context === undefined) {
        throw new Error('useAuth must be used within an AuthProvider');
    }
    return context;
};