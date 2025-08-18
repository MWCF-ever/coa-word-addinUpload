import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { Stack, Text, MessageBar, MessageBarType, Spinner, SpinnerSize, PrimaryButton, DefaultButton, ProgressIndicator } from '@fluentui/react';
import { CompoundSelector } from './CompoundSelector';
import { TemplateSelector } from './TemplateSelector';
import { AuthGuard, UserInfo } from '../../components/AuthGuard';
import { useAuth } from '../../contexts/AuthContext';
import { AppState, Compound, Template, BatchData } from '../../types';
import { apiClient } from '../../services/httpInterceptor';
// 导入新的批次模板服务
import { createAIMTADocument } from '../../services/batchTemplete';
import '../taskpane.css';

// 连接状态类型
interface ConnectionStatus {
    isConnected: boolean;
    message: string;
    apiUrl: string;
    lastChecked?: Date;
    responseTime?: number;
}

// 连接测试函数
const testApiConnection = async (): Promise<{ success: boolean; responseTime: number; error?: string }> => {
    const startTime = Date.now();
    try {
        console.log(`🔍 测试API连接: ${apiClient.defaults.baseURL}`);
        
        const response = await apiClient.get('/health');
        
        const responseTime = Date.now() - startTime;
        
        if (response.status === 200) {
            console.log(`✅ API连接成功 (${responseTime}ms):`, response.data);
            return { success: true, responseTime };
        } else {
            console.error(`❌ API连接失败: ${response.status} ${response.statusText}`);
            return { success: false, responseTime, error: `HTTP ${response.status}: ${response.statusText}` };
        }
    } catch (error: any) {
        const responseTime = Date.now() - startTime;
        console.error(`❌ API连接异常:`, error);
        
        let errorMessage = 'Unknown error';
        if (error.response) {
            errorMessage = `HTTP ${error.response.status}: ${error.response.statusText}`;
        } else if (error.request) {
            errorMessage = 'Network connection failed';
        } else {
            errorMessage = error.message;
        }
        
        return { success: false, responseTime, error: errorMessage };
    }
};

export const App: React.FC = () => {
    // 使用认证上下文
    const { user, isAuthenticated, getToken } = useAuth();
    
    // 使用useRef来存储不需要触发重新渲染的值
    const mountedRef = useRef(true);
    const processingRef = useRef(false);

    // 合并相关状态，减少状态更新次数
    const [state, setState] = useState<AppState>({
        extractedData: [],
        isLoading: false,
        selectedCompound: undefined,
        selectedTemplate: undefined,
        error: undefined
    });

    const [compounds, setCompounds] = useState<Compound[]>([]);
    const [templates, setTemplates] = useState<Template[]>([]);
    const [batchDataList, setBatchDataList] = useState<BatchData[]>([]);
    const [processingStatus, setProcessingStatus] = useState<string>('');
    
    // 连接状态
    const [connectionStatus, setConnectionStatus] = useState<ConnectionStatus>({
        isConnected: false,
        message: '正在检查API连接...',
        apiUrl: apiClient.defaults.baseURL || ''
    });

    // 组件卸载时清理
    useEffect(() => {
        return () => {
            mountedRef.current = false;
        };
    }, []);

    // 连接测试函数
    const checkConnection = useCallback(async (showLoading: boolean = true) => {
        if (showLoading) {
            setState(prev => ({ ...prev, isLoading: true }));
        }
        
        setConnectionStatus(prev => ({
            ...prev,
            message: '正在测试连接...'
        }));
        
        const result = await testApiConnection();
        
        setConnectionStatus({
            isConnected: result.success,
            message: result.success 
                ? `✅ 连接成功 (${result.responseTime}ms)` 
                : `❌ 连接失败: ${result.error}`,
            apiUrl: apiClient.defaults.baseURL || '',
            lastChecked: new Date(),
            responseTime: result.responseTime
        });
        
        if (showLoading) {
            setState(prev => ({ ...prev, isLoading: false }));
        }
        
        return result.success;
    }, []);

    // 应用启动时检查连接（只有在已认证时才执行）
    useEffect(() => {
        const initializeApp = async () => {
            if (!isAuthenticated) {
                return;
            }

            console.log('🚀 AIMTA应用初始化开始!!');
            console.log(`📡 API地址: ${apiClient.defaults.baseURL}`);
            console.log(`🌐 当前域名: ${window.location.hostname}`);
            console.log(`🔗 完整URL: ${window.location.href}`);
            console.log(`👤 当前用户: ${user?.name} (${user?.email})`);
            
            const isConnected = await checkConnection(false);
            
            if (isConnected) {
                fetchCompounds();
            } else {
                setState(prev => ({ 
                    ...prev, 
                    error: `后端API连接失败。请检查:\n1. 网络连接\n2. 后端服务是否正常运行\n3. API地址是否正确: ${apiClient.defaults.baseURL}`
                }));
            }
        };
        
        if (isAuthenticated) {
            initializeApp();
        }
    }, [isAuthenticated, user, checkConnection]);

    // 获取化合物列表
    const fetchCompounds = useCallback(async () => {
        if (!mountedRef.current || !isAuthenticated) return;
        
        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            console.log(`📡 请求化合物列表: ${apiClient.defaults.baseURL}/compounds`);
            
            const response = await apiClient.get('/compounds');
            
            if (mountedRef.current) {
                console.log('✅ 化合物数据获取成功:', response.data);
                setCompounds(response.data.data || []);
            }
        } catch (error: any) {
            console.error('❌ 获取化合物列表失败:', error);
            
            if (mountedRef.current) {
                let errorMessage = '获取化合物列表失败。';
                if (error.response) {
                    if (error.response.status === 401) {
                        errorMessage += ' 认证失败，请重新登录。';
                    } else if (error.response.status === 403) {
                        errorMessage += ' 访问被拒绝，请检查权限。';
                    } else {
                        errorMessage += ` HTTP ${error.response.status}: ${error.response.statusText}`;
                    }
                } else if (error.request) {
                    errorMessage += ' 网络连接失败。';
                } else {
                    errorMessage += ` ${error.message}`;
                }
                
                setState(prev => ({ 
                    ...prev, 
                    error: errorMessage
                }));
                
                setConnectionStatus(prev => ({
                    ...prev,
                    isConnected: false,
                    message: '❌ API连接已断开',
                    lastChecked: new Date()
                }));
            }
        } finally {
            if (mountedRef.current) {
                setState(prev => ({ ...prev, isLoading: false }));
            }
        }
    }, [isAuthenticated]);

    // 获取模板列表
    const fetchTemplates = useCallback(async (compoundId: string) => {
        if (!mountedRef.current || !isAuthenticated) return;
        
        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            console.log(`📡 请求模板列表: compound_id=${compoundId}`);
            
            const response = await apiClient.get('/templates', {
                params: { compound_id: compoundId }
            });
            
            if (mountedRef.current) {
                console.log('✅ 模板数据获取成功:', response.data);
                setTemplates(response.data.data || []);
            }
        } catch (error) {
            console.error('❌ 获取模板列表失败:', error);
            if (mountedRef.current) {
                setState(prev => ({ 
                    ...prev, 
                    error: '获取模板列表失败。请重试。' 
                }));
            }
        } finally {
            if (mountedRef.current) {
                setState(prev => ({ ...prev, isLoading: false }));
            }
        }
    }, [isAuthenticated]);

    // 当化合物变化时获取模板
    useEffect(() => {
        if (state.selectedCompound && mountedRef.current && isAuthenticated) {
            fetchTemplates(state.selectedCompound.id);
        }
    }, [state.selectedCompound, fetchTemplates, isAuthenticated]);

    // 处理化合物选择
    const handleCompoundSelect = useCallback((compound: Compound) => {
        console.log('📋 选择化合物:', compound);
        setState(prev => ({ 
            ...prev, 
            selectedCompound: compound,
            selectedTemplate: undefined,
            extractedData: []
        }));
        setBatchDataList([]);
        setProcessingStatus('');
    }, []);

    // 处理模板选择
    const handleTemplateSelect = useCallback((template: Template) => {
        console.log('📋 选择模板:', template);
        setState(prev => ({ 
            ...prev, 
            selectedTemplate: template 
        }));
        setProcessingStatus('');
    }, []);

    // 重新连接处理
    const handleReconnect = useCallback(async () => {
        const isConnected = await checkConnection();
        if (isConnected) {
            fetchCompounds();
        }
    }, [checkConnection, fetchCompounds]);

    // 处理文件处理
    const handleProcessFiles = useCallback(async () => {
        if (!state.selectedCompound || !state.selectedTemplate || !isAuthenticated) {
            setState(prev => ({ 
                ...prev, 
                error: 'Please select compound and template first.' 
            }));
            return;
        }

        if (processingRef.current) return;
        processingRef.current = true;

        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            setProcessingStatus('Checking database for existing batch analysis data...');
            
            console.log('🔍 发送API请求:', {
                compound_id: state.selectedCompound.id,
                template_id: state.selectedTemplate.id
            });
            
            const cacheCheckResponse = await apiClient.get('/documents/check-cache', { 
                params: {
                    compound_id: state.selectedCompound.id,
                    template_id: state.selectedTemplate.id
                }
            });

            console.log('📡 完整API响应:', cacheCheckResponse);
            
            const cachedData = cacheCheckResponse.data.data;
            console.log('🗂️ 提取的缓存数据:', cachedData);
            
            if (cachedData && cachedData.batchData && cachedData.batchData.length > 0) {
                const batchData = cachedData.batchData;
                console.log('🎯 准备设置批次数据:', batchData);
                
                setBatchDataList(batchData);
                console.log('✅ setBatchDataList 调用完成');
                
                setProcessingStatus(
                    `Found existing data: ${batchData.length} batches loaded from database. ` +
                    `Last updated: ${new Date(cachedData.lastUpdated).toLocaleString()}`
                );
                
                setState(prev => ({ ...prev, isLoading: false }));
                
            } else {
                console.log('🔭 没有缓存数据，开始处理PDF文件');
                
                setProcessingStatus('No cached data found. Scanning PDF files and extracting batch analysis data...');
                
                const requestData = {
                    compound_id: state.selectedCompound.id,
                    template_id: state.selectedTemplate.id,
                    force_reprocess: false
                };
                
                console.log('📤 发送处理请求:', requestData);
                
                const processResponse = await apiClient.post('/documents/process-directory', requestData);
                console.log('📥 处理响应:', processResponse.data);

                const batchData = processResponse.data.data.batchData || [];
                console.log('🎯 新处理的批次数据:', batchData);
                
                setBatchDataList(batchData);
                setProcessingStatus(`Successfully processed and cached ${batchData.length} batches!`);
                setState(prev => ({ ...prev, isLoading: false }));
            }

        } catch (error: any) {
            console.error('❌ 处理文件时出错:', error);
            
            let errorMessage = 'Failed to process files. ';
            if (error.response) {
                if (error.response.status === 401) {
                    errorMessage += 'Authentication failed. Please log in again.';
                } else if (error.response.status === 403) {
                    errorMessage += 'Access denied. Please check your permissions.';
                } else if (error.response.status === 422) {
                    errorMessage += 'Request validation failed. ';
                    if (error.response.data?.detail) {
                        errorMessage += `Details: ${JSON.stringify(error.response.data.detail)}`;
                    }
                } else if (error.response.data?.detail) {
                    errorMessage += error.response.data.detail;
                } else {
                    errorMessage += 'Please try again.';
                }
            } else {
                errorMessage += 'Please try again.';
            }
            
            setState(prev => ({ 
                ...prev, 
                error: errorMessage,
                isLoading: false
            }));
            setProcessingStatus('');
            
        } finally {
            processingRef.current = false;
            console.log('🏁 handleProcessFiles 执行完成');
        }
    }, [state.selectedCompound, state.selectedTemplate, isAuthenticated]);

    // 创建Word文档 - 使用新的服务
    const handleCreateWordDocument = useCallback(async () => {
        if (!state.selectedTemplate || batchDataList.length === 0) {
            setState(prev => ({ 
                ...prev, 
                error: 'No batch analysis data available to create document.' 
            }));
            return;
        }

        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            setProcessingStatus('Creating complete AIMTA document with extracted batch data...');
            
            // 使用新的批次模板服务
            const compoundCode = state.selectedCompound?.code || 'BGB-16673';
            await createAIMTADocument(batchDataList, compoundCode);
            
            setProcessingStatus('Complete AIMTA Document created successfully with headers, footers and template formatting!');
            
        } catch (error) {
            console.error('Failed to create document:', error);
            setState(prev => ({ 
                ...prev, 
                error: 'Failed to create Word document with template formatting.' 
            }));
        } finally {
            setState(prev => ({ ...prev, isLoading: false }));
        }
    }, [state.selectedTemplate, state.selectedCompound, batchDataList]);

    // 调试输出
    useEffect(() => {
        console.log('🔍 =================== 渲染状态调试 ===================');
        console.log('⏰ 时间戳:', new Date().toLocaleTimeString());
        console.log('📊 batchDataList.length:', batchDataList.length);
        console.log('📋 batchDataList内容:', batchDataList);
        console.log('🔄 state.isLoading:', state.isLoading);
        console.log('🧪 selectedCompound:', state.selectedCompound?.code);
        console.log('📄 selectedTemplate:', state.selectedTemplate?.region);
        console.log('✅ 按钮组应该显示:', batchDataList.length > 0);
        console.log('📝 processingStatus:', processingStatus);
        console.log('🔍 ======================================================');
    }, [batchDataList, state.isLoading, state.selectedCompound, state.selectedTemplate, processingStatus]);

    return (
        <AuthGuard requireAuth={true}>
            <Stack className="app-container" tokens={{ childrenGap: 20 }}>
                <Stack.Item>
                    <Text variant="xLarge" className="app-title">AIMTA Batch Analysis Processor</Text>
                </Stack.Item>

                {/* 用户信息显示 */}
                <UserInfo />

                {/* 连接状态显示 */}
                <MessageBar 
                    messageBarType={connectionStatus.isConnected ? MessageBarType.success : MessageBarType.warning}
                    isMultiline={true}
                >
                    <div>
                        <strong>API连接状态:</strong> {connectionStatus.message}
                        <br />
                        <small>API地址: {connectionStatus.apiUrl}</small>
                        {connectionStatus.lastChecked && (
                            <>
                                <br />
                                <small>最后检查: {connectionStatus.lastChecked.toLocaleTimeString()}</small>
                            </>
                        )}
                        {connectionStatus.responseTime && (
                            <>
                                <br />
                                <small>响应时间: {connectionStatus.responseTime}ms</small>
                            </>
                        )}
                    </div>
                </MessageBar>

                {state.error && (
                    <MessageBar 
                        messageBarType={MessageBarType.error}
                        onDismiss={() => setState(prev => ({ ...prev, error: undefined }))}
                        isMultiline={true}
                    >
                        {state.error}
                    </MessageBar>
                )}

                {state.isLoading && (
                    <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
                        <Spinner size={SpinnerSize.large} label="Processing..." />
                    </Stack>
                )}

                {processingStatus && (
                    <MessageBar messageBarType={MessageBarType.info}>
                        {processingStatus}
                    </MessageBar>
                )}

                {/* 如果连接失败，显示重试按钮 */}
                {!connectionStatus.isConnected && (
                    <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
                        <PrimaryButton
                            text="重新测试连接"
                            iconProps={{ iconName: 'Refresh' }}
                            onClick={handleReconnect}
                            disabled={state.isLoading}
                        />
                    </Stack>
                )}

                <Stack tokens={{ childrenGap: 15 }}>
                    <CompoundSelector
                        compounds={compounds}
                        selectedCompound={state.selectedCompound}
                        onSelect={handleCompoundSelect}
                        disabled={state.isLoading}
                    />

                    {state.selectedCompound && (
                        <TemplateSelector
                            templates={templates}
                            selectedTemplate={state.selectedTemplate}
                            onSelect={handleTemplateSelect}
                            disabled={state.isLoading}
                        />
                    )}

                    {state.selectedCompound && state.selectedTemplate && (
                        <Stack tokens={{ childrenGap: 10 }}>
                            <PrimaryButton
                                text="Load/Process AIMTA Files"
                                iconProps={{ iconName: 'DatabaseSearch' }}
                                onClick={handleProcessFiles}
                                disabled={state.isLoading}
                                styles={{ root: { width: '100%' } }}
                            />

                            {batchDataList.length > 0 && (
                                <Stack tokens={{ childrenGap: 10 }}>
                                    <MessageBar messageBarType={MessageBarType.success}>
                                        {`Ready to create document: ${batchDataList.length} batches available`}
                                        <br />
                                        {`Batch data: ${batchDataList.map(b => b.batch_number).join(', ')}`}
                                    </MessageBar>
                                    
                                    {/* 主要文档创建按钮 */}
                                    <PrimaryButton
                                        text="Create Complete AIMTA Document"
                                        iconProps={{ iconName: 'NewTeamProject' }}
                                        onClick={handleCreateWordDocument}
                                        disabled={state.isLoading}
                                        styles={{ 
                                            root: { 
                                                width: '100%',
                                                backgroundColor: '#0078d4',
                                                borderColor: '#0078d4'
                                            } 
                                        }}
                                    />
                                </Stack>
                            )}
                        </Stack>
                    )}
                </Stack>

                <Stack className="help-text">
                    <Text variant="small">  - 🔒 Secure authentication with BeiGene SSO</Text>
                    <Text variant="small">  - Select compound and template region</Text>
                    <Text variant="small">  - Click "Load/Process AIMTA Files" to check database first, then process if needed</Text>
                    <Text variant="small">  - "Create Complete AIMTA Document" generates a full formatted document</Text>
                    <Text variant="small">  - All API requests are automatically authenticated</Text>
                    <Text variant="small">  - Token refresh happens automatically every 5 minutes</Text>
                </Stack>
            </Stack>
        </AuthGuard>
    );
};