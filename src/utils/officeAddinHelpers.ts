// src/utils/officeAddinHelpers.ts - Office加载项辅助函数
export const isOfficeAddinEnvironment = (): boolean => {
    const hostname = window.location.hostname;
    const userAgent = navigator.userAgent.toLowerCase();
    
    return (
        hostname.includes('sharepoint.com') ||
        hostname.includes('office.com') ||
        hostname.includes('officeapps.live.com') ||
        userAgent.includes('office') ||
        userAgent.includes('microsoft') ||
        // 检查是否在Office.js环境中
        (typeof Office !== 'undefined' && !!Office.context)
    );
};

// Office加载项专用连接测试
export const testOfficeAddinConnection = async (): Promise<{
    success: boolean;
    responseTime: number;
    error?: string;
    officeEnvironment?: any;
}> => {
    const startTime = Date.now();
    
    try {
        console.log('🏢 开始Office加载项连接测试...');
        
        // 动态导入apiClient以避免循环依赖
        const { apiClient } = await import('../services/httpInterceptor');
        
        // 测试Office加载项专用端点
        let healthResult;
        try {
            const response = await apiClient.get('/office-addin-health');
            healthResult = response.data;
        } catch (error) {
            console.warn('Office专用端点失败，尝试标准端点');
            const response = await apiClient.get('/health');
            healthResult = response.data;
        }
        
        const responseTime = Date.now() - startTime;
        
        console.log('✅ Office加载项连接测试成功:', healthResult);
        
        // 额外获取Office环境信息
        let officeEnvironment = null;
        try {
            const response = await apiClient.get('/detect-office-addin');
            officeEnvironment = response.data;
        } catch (e) {
            console.warn('获取Office环境信息失败:', e);
        }
        
        return {
            success: true,
            responseTime,
            officeEnvironment
        };
        
    } catch (error: any) {
        const responseTime = Date.now() - startTime;
        console.error('❌ Office加载项连接测试失败:', error);
        
        let errorMessage = 'Unknown error';
        if (error.response) {
            errorMessage = `HTTP ${error.response.status}: ${error.response.statusText}`;
        } else if (error.request) {
            errorMessage = 'Network connection failed - possible CORS issue';
        } else {
            errorMessage = error.message;
        }
        
        return {
            success: false,
            responseTime,
            error: errorMessage
        };
    }
};

// 完整的Office加载项测试套件
export const fullOfficeAddinTest = async () => {
    console.log('🏢 开始Office加载项完整测试...');
    
    // 1. 环境检测
    const envInfo = {
        isOffice: isOfficeAddinEnvironment(),
        hostname: window.location.hostname,
        hasOfficeJs: typeof Office !== 'undefined',
        origin: window.location.origin,
        userAgent: navigator.userAgent.substring(0, 100)
    };
    
    console.log('环境检测:', envInfo);
    
    // 2. 连接测试
    const connectionResult = await testOfficeAddinConnection();
    console.log('连接测试结果:', connectionResult);
    
    // 3. 专用端点测试
    try {
        const { apiClient } = await import('../services/httpInterceptor');
        
        const endpointTests = [];
        
        // 测试各个端点
        const endpoints = [
            { path: '/office-addin-health', name: 'Office专用健康检查' },
            { path: '/cors-test', name: 'CORS测试' },
            { path: '/detect-office-addin', name: 'Office环境检测' },
            { path: '/auth/status', name: '认证状态检查' },
            { path: '/health', name: '标准健康检查' }
        ];
        
        for (const endpoint of endpoints) {
            try {
                const startTime = Date.now();
                const response = await apiClient.get(endpoint.path);
                const responseTime = Date.now() - startTime;
                
                endpointTests.push({
                    name: endpoint.name,
                    path: endpoint.path,
                    success: true,
                    status: response.status,
                    responseTime,
                    data: response.data
                });
                
                console.log(`✅ ${endpoint.name} 成功:`, response.data);
            } catch (error: any) {
                endpointTests.push({
                    name: endpoint.name,
                    path: endpoint.path,
                    success: false,
                    error: error.message,
                    status: error.response?.status
                });
                
                console.error(`❌ ${endpoint.name} 失败:`, error.message);
            }
        }
        
        const successCount = endpointTests.filter(t => t.success).length;
        const totalCount = endpointTests.length;
        
        console.log(`📊 端点测试汇总: ${successCount}/${totalCount} 成功`);
        console.log('详细结果:', endpointTests);
        
        return { 
            success: successCount > 0, 
            message: `Office加载项测试完成: ${successCount}/${totalCount} 端点成功`,
            envInfo,
            connectionResult,
            endpointTests
        };
    } catch (error: any) {
        console.error('❌ Office加载项测试失败:', error);
        return { success: false, error: error.message, envInfo, connectionResult };
    }
};

// 暴露到全局window对象
declare global {
    interface Window {
        isOfficeAddinEnvironment: () => boolean;
        testOfficeAddinConnection: () => Promise<any>;
        fullOfficeAddinTest: () => Promise<any>;
        officeAddinApiClient: any;
    }
}

// 立即暴露函数到全局
if (typeof window !== 'undefined') {
    window.isOfficeAddinEnvironment = isOfficeAddinEnvironment;
    window.testOfficeAddinConnection = testOfficeAddinConnection;
    window.fullOfficeAddinTest = fullOfficeAddinTest;
    
    // 延迟暴露apiClient，等待模块加载完成
    setTimeout(async () => {
        try {
            const { apiClient } = await import('../services/httpInterceptor');
            window.officeAddinApiClient = apiClient;
            console.log('🌐 Office加载项调试工具已加载到window对象');
        } catch (error) {
            console.warn('⚠️ 无法加载apiClient到全局:', error);
        }
    }, 1000);
}