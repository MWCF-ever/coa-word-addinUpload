import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { Stack, Text, MessageBar, MessageBarType, Spinner, SpinnerSize, PrimaryButton, DefaultButton, ProgressIndicator } from '@fluentui/react';
import { CompoundSelector } from './CompoundSelector';
import { TemplateSelector } from './TemplateSelector';
import { BatchDataTable } from './BatchDataTable';
import { useAsyncProcessing } from '../hooks/useAsyncProcessing';
import { useWordDocument } from '../hooks/useWordDocument';
import { AppState, Compound, Template, BatchData, API_BASE_URL } from '../../types';
import axios from 'axios';
import '../taskpane.css';

// 创建axios实例，设置默认配置
const api = axios.create({
    baseURL: API_BASE_URL,
    timeout: 30000,
});

// 添加连接状态类型
interface ConnectionStatus {
    isConnected: boolean;
    message: string;
    apiUrl: string;
    lastChecked?: Date;
    responseTime?: number;
}

// 添加连接测试函数
const testApiConnection = async (): Promise<{ success: boolean; responseTime: number; error?: string }> => {
    const startTime = Date.now();
    try {
        console.log(`🔍 测试API连接: ${API_BASE_URL}`);
        
        const response = await axios.get(`${API_BASE_URL}/api/health`, {
            timeout: 10000,
            headers: {
                'Content-Type': 'application/json',
            },
        });
        
        const responseTime = Date.now() - startTime;
        
        if (response.status === 200) {
            console.log(`✅ API连接成功 (${responseTime}ms):`, response.data);
            return { success: true, responseTime };
        } else {
            console.error(`❌ API连接失败: ${response.status} ${response.statusText}`);
            return { success: false, responseTime, error: `HTTP ${response.status}: ${response.statusText}` };
        }
    } catch (error) {
        const responseTime = Date.now() - startTime;
        console.error(`❌ API连接异常:`, error);
        
        let errorMessage = 'Unknown error';
        if (axios.isAxiosError(error)) {
            if (error.code === 'ECONNABORTED') {
                errorMessage = 'Request timeout';
            } else if (error.response) {
                errorMessage = `HTTP ${error.response.status}: ${error.response.statusText}`;
            } else if (error.request) {
                errorMessage = 'Network connection failed';
            } else {
                errorMessage = error.message;
            }
        }
        
        return { success: false, responseTime, error: errorMessage };
    }
};

export const App: React.FC = () => {
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
    
    // 添加连接状态
    const [connectionStatus, setConnectionStatus] = useState<ConnectionStatus>({
        isConnected: false,
        message: '正在检查API连接...',
        apiUrl: API_BASE_URL
    });

    // 组件卸载时清理
    useEffect(() => {
        return () => {
            mountedRef.current = false;
        };
    }, []);

    // 添加连接测试函数
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
            apiUrl: API_BASE_URL,
            lastChecked: new Date(),
            responseTime: result.responseTime
        });
        
        if (showLoading) {
            setState(prev => ({ ...prev, isLoading: false }));
        }
        
        return result.success;
    }, []);

    // 应用启动时检查连接
    useEffect(() => {
        const initializeApp = async () => {
            console.log('🚀 AIMTA应用初始化开始');
            console.log(`📡 API地址: ${API_BASE_URL}`);
            console.log(`🌍 当前域名: ${window.location.hostname}`);
            console.log(`🔗 完整URL: ${window.location.href}`);
            
            // 检查连接
            const isConnected = await checkConnection(false);
            
            if (isConnected) {
                // 连接成功，获取化合物数据
                fetchCompounds();
            } else {
                // 连接失败，显示错误信息
                setState(prev => ({ 
                    ...prev, 
                    error: `后端API连接失败。请检查:\n1. 网络连接\n2. 后端服务是否正常运行\n3. API地址是否正确: ${API_BASE_URL}`
                }));
            }
        };
        
        initializeApp();
    }, [checkConnection]);

    // 使用useCallback优化回调函数
    const fetchCompounds = useCallback(async () => {
        if (!mountedRef.current) return;
        
        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            console.log(`📡 请求化合物列表: ${API_BASE_URL}/api/compounds`);
            
            const response = await axios.get(`${API_BASE_URL}/api/compounds`, {
                timeout: 10000,
            });
            
            if (mountedRef.current) {
                console.log('✅ 化合物数据获取成功:', response.data);
                setCompounds(response.data.data || []);
            }
        } catch (error) {
            console.error('❌ 获取化合物列表失败:', error);
            
            if (mountedRef.current) {
                let errorMessage = '获取化合物列表失败。';
                if (axios.isAxiosError(error)) {
                    if (error.code === 'ECONNABORTED') {
                        errorMessage += ' 请求超时。';
                    } else if (error.response) {
                        errorMessage += ` HTTP ${error.response.status}: ${error.response.statusText}`;
                    } else if (error.request) {
                        errorMessage += ' 网络连接失败。';
                    } else {
                        errorMessage += ` ${error.message}`;
                    }
                }
                
                setState(prev => ({ 
                    ...prev, 
                    error: errorMessage
                }));
                
                // 连接失败时更新连接状态
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
    }, []);

    const fetchTemplates = useCallback(async (compoundId: string) => {
        if (!mountedRef.current) return;
        
        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            console.log(`📡 请求模板列表: compound_id=${compoundId}`);
            
            const response = await axios.get(`${API_BASE_URL}/api/templates`, {
                params: { compound_id: compoundId },
                timeout: 10000,
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
    }, []);

    // Fetch templates when compound changes
    useEffect(() => {
        if (state.selectedCompound && mountedRef.current) {
            fetchTemplates(state.selectedCompound.id);
        }
    }, [state.selectedCompound, fetchTemplates]);

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

    const handleTemplateSelect = useCallback((template: Template) => {
        console.log('📋 选择模板:', template);
        setState(prev => ({ 
            ...prev, 
            selectedTemplate: template 
        }));
        setProcessingStatus('');
    }, []);

    // 添加重新连接按钮处理函数
    const handleReconnect = useCallback(async () => {
        const isConnected = await checkConnection();
        if (isConnected) {
            // 重新获取数据
            fetchCompounds();
        }
    }, [checkConnection, fetchCompounds]);

    const handleProcessFiles = useCallback(async () => {
        if (!state.selectedCompound || !state.selectedTemplate) {
            setState(prev => ({ 
                ...prev, 
                error: 'Please select compound and template first.' 
            }));
            return;
        }

        // 防止重复点击
        if (processingRef.current) return;
        processingRef.current = true;

        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            setProcessingStatus('Checking database for existing batch analysis data...');
            
            // 第一步：检查数据库中是否已有数据
            const cacheCheckResponse = await axios.get(
                `${API_BASE_URL}/api/documents/check-cache`,
                { 
                    params: {
                        compound_id: state.selectedCompound.id,
                        template_id: state.selectedTemplate.id
                    }
                }
            );

            const cachedData = cacheCheckResponse.data.data;
            
            if (cachedData && cachedData.batchData && cachedData.batchData.length > 0) {
                // 如果数据库中有数据，直接使用
                setBatchDataList(cachedData.batchData);
                setProcessingStatus(
                    `Found existing data: ${cachedData.batchData.length} batches loaded from database. ` +
                    `Last updated: ${new Date(cachedData.lastUpdated).toLocaleString()}`
                );
                
                console.log('使用缓存数据:', cachedData.batchData);
            } else {
                // 如果数据库中没有数据，进行PDF处理
                setProcessingStatus('No cached data found. Scanning PDF files and extracting batch analysis data...');
                
                const requestData = {
                    compound_id: state.selectedCompound.id,
                    template_id: state.selectedTemplate.id,
                    force_reprocess: false
                };
                
                console.log('发送处理请求:', requestData);
                
                const processResponse = await axios.post(
                    `${API_BASE_URL}/api/documents/process-directory`,
                    requestData,
                    {
                        headers: {
                            'Content-Type': 'application/json'
                        }
                    }
                );

                const batchData = processResponse.data.data.batchData || [];
                setBatchDataList(batchData);
                
                setProcessingStatus(`Successfully processed and cached ${batchData.length} batches!`);
                console.log('新处理数据:', batchData);
            }

        } catch (error) {
            console.error('处理文件时出错:', error);
            
            if (axios.isAxiosError(error) && error.response) {
                const response = error.response;
                console.error('响应错误:', response.data);
                console.error('状态码:', response.status);
                
                let errorMessage = 'Failed to process files. ';
                if (response.status === 422) {
                    errorMessage += 'Request validation failed. ';
                    if (response.data?.detail) {
                        errorMessage += `Details: ${JSON.stringify(response.data.detail)}`;
                    }
                } else if (response.data?.detail) {
                    errorMessage += response.data.detail;
                } else {
                    errorMessage += 'Please try again.';
                }
                
                setState(prev => ({ 
                    ...prev, 
                    error: errorMessage
                }));
            } else {
                setState(prev => ({ 
                    ...prev, 
                    error: 'Failed to process files. Please try again.' 
                }));
            }
            setProcessingStatus('');
        } finally {
            processingRef.current = false;
            if (mountedRef.current) {
                setState(prev => ({ ...prev, isLoading: false }));
            }
        }
    }, [state.selectedCompound, state.selectedTemplate]);


    const handleForceReprocess = useCallback(async () => {
        if (!state.selectedCompound || !state.selectedTemplate) {
            setState(prev => ({ 
                ...prev, 
                error: 'Please select compound and template first.' 
            }));
            return;
        }

        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            setProcessingStatus('Force reprocessing: Scanning PDF files and updating database...');
            
            const requestData = {
                compound_id: state.selectedCompound.id,
                template_id: state.selectedTemplate.id,
                force_reprocess: true
            };
            
            const processResponse = await axios.post(
                `${API_BASE_URL}/api/documents/process-directory`,
                requestData,
                {
                    headers: {
                        'Content-Type': 'application/json'
                    }
                }
            );

            const batchData = processResponse.data.data.batchData || [];
            setBatchDataList(batchData);
            
            setProcessingStatus(`Force reprocessed and updated ${batchData.length} batches in database!`);
            console.log('强制重新处理数据:', batchData);

        } catch (error) {
            console.error('强制重新处理时出错:', error);
            setState(prev => ({ 
                ...prev, 
                error: 'Failed to force reprocess files. Please try again.' 
            }));
            setProcessingStatus('');
        } finally {
            setState(prev => ({ ...prev, isLoading: false }));
        }
    }, [state.selectedCompound, state.selectedTemplate]);

    const handleClearCache = useCallback(async () => {
        if (!state.selectedCompound || !state.selectedTemplate) {
            setState(prev => ({ 
                ...prev, 
                error: 'Please select compound and template first.' 
            }));
            return;
        }

        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            setProcessingStatus('Clearing cached data...');
            
            await axios.delete(
                `${API_BASE_URL}/api/documents/clear-cache`,
                { 
                    params: {
                        compound_id: state.selectedCompound.id,
                        template_id: state.selectedTemplate.id
                    }
                }
            );

            setBatchDataList([]);
            setProcessingStatus('Cache cleared successfully!');
            console.log('缓存已清除');

        } catch (error) {
            console.error('清除缓存时出错:', error);
            setState(prev => ({ 
                ...prev, 
                error: 'Failed to clear cache. Please try again.' 
            }));
            setProcessingStatus('');
        } finally {
            setState(prev => ({ ...prev, isLoading: false }));
        }
    }, [state.selectedCompound, state.selectedTemplate]);

    // 根据模板信息创建完整的Word文档 - 优化版本
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
            
            await Word.run(async (context) => {
                // 清空文档
                context.document.body.clear();
                
                // 批量操作：收集所有操作后统一执行
                const operations: (() => void)[] = [];
                
                // 创建目录和表格列表（在同一页）
                operations.push(() => createTableOfContentsAndListOfTables(context));
                
                // 插入分页符
                operations.push(() => {
                    context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
                });
                
                // 创建主要内容 - S.4.4 Batch Analyses
                operations.push(() => createBatchAnalysesSection(context));
                
                // 执行所有操作
                operations.forEach(op => op());
                
                // 只在最后执行一次sync
                await context.sync();
            });
            
            setProcessingStatus('Complete AIMTA Document created successfully with template formatting!');
            
        } catch (error) {
            console.error('Failed to create document:', error);
            setState(prev => ({ 
                ...prev, 
                error: 'Failed to create Word document with template formatting.' 
            }));
        } finally {
            setState(prev => ({ ...prev, isLoading: false }));
        }
    }, [state.selectedTemplate, batchDataList]);


    // 根据模板信息创建完整文档结构
    const createTableOfContentsAndListOfTables = (context: Word.RequestContext) => {
        try {
            // 标题: Table of Contents
            const tocTitle = context.document.body.insertParagraph('Table of Contents', Word.InsertLocation.end);
            tocTitle.alignment = Word.Alignment.centered;
            tocTitle.font.name = 'Times New Roman';
            tocTitle.font.size = 14;
            tocTitle.font.bold = true;
            tocTitle.font.color = '#000000';
            tocTitle.spaceAfter = 12;
            tocTitle.spaceBefore = 0;

            // TOC 条目
            const tocItems = [
                { text: 'Table of Contents', page: '1' },
                { text: 'List of Tables', page: '1' },
                { text: 'S.4.4\tBatch Analyses', page: '2' }
            ];

            tocItems.forEach(item => {
                const tocEntry = context.document.body.insertParagraph('', Word.InsertLocation.end);
                
                // 插入带有内部链接的文本
                if (item.text === 'S.4.4\tBatch Analyses') {
                    tocEntry.insertText(item.text, Word.InsertLocation.end);
                    tocEntry.insertText('\t' + item.page, Word.InsertLocation.end);
                    // 注意：Word JS API 目前不支持直接创建内部超链接
                    // 需要使用书签或其他方法实现
                } else {
                    tocEntry.insertText(item.text + '\t' + item.page, Word.InsertLocation.end);
                }
                
                tocEntry.alignment = Word.Alignment.left;
                tocEntry.font.name = 'Times New Roman';
                tocEntry.font.size = 12;
                tocEntry.font.color = '#0000FF';
                tocEntry.leftIndent = 0;
                tocEntry.spaceAfter = 0;
                tocEntry.spaceBefore = 0;
            });

            // 空行分隔
            context.document.body.insertParagraph('', Word.InsertLocation.end);

            // 标题: List of Tables - 在同一页
            const listTitle = context.document.body.insertParagraph('List of Tables', Word.InsertLocation.end);
            listTitle.alignment = Word.Alignment.centered;
            listTitle.font.name = 'Times New Roman';
            listTitle.font.size = 14;
            listTitle.font.bold = true;
            listTitle.font.color = '#000000';
            listTitle.spaceAfter = 12;
            listTitle.spaceBefore = 12;

            // 表格列表条目
            const tableItems = [
                { text: 'Table 1:\tOverview of BGB-16673 Drug Substance Batches', page: '2' },
                { text: 'Table 2:\tBatch Analysis for Toxicology Batches of BGB-16673 Drug Substance', page: '3' },
                { text: 'Table 3:\tBatch Analysis for GMP Batches of BGB-16673 Drug Substance', page: '5' },
                { text: 'Table 4:\tBatch Results for GMP Batches of BGB-16673 Drug Substance', page: '7' }
            ];

            tableItems.forEach(item => {
                const tableEntry = context.document.body.insertParagraph(item.text + '\t' + item.page, Word.InsertLocation.end);
                tableEntry.alignment = Word.Alignment.left;
                tableEntry.font.name = 'Times New Roman';
                tableEntry.font.size = 12;
                tableEntry.font.color = '#0000FF';
                tableEntry.leftIndent = 0;
                tableEntry.spaceAfter = 0;
                tableEntry.spaceBefore = 0;
            });
        } catch (error) {
            console.error('创建目录时出错:', error);
            throw error;
        }
    };


    // 创建批次分析主要内容 - 应用模板格式
    const createBatchAnalysesSection = async (context: Word.RequestContext) => {
        try {
            // S.4.4 标题 - 应用C-Heading 1 (non-numbered)样式
            const sectionTitle = context.document.body.insertParagraph('S.4.4	BATCH ANALYSES', Word.InsertLocation.end);
            sectionTitle.font.name = 'Times New Roman';
            sectionTitle.font.size = 14;
            sectionTitle.font.bold = true;
            sectionTitle.font.color = '#000000';
            sectionTitle.spaceAfter = 12;
            sectionTitle.spaceBefore = 0;
            sectionTitle.alignment = Word.Alignment.left;
            sectionTitle.font.italic = false;

            // 介绍段落 - 应用C-Body Text样式
            const introPara = context.document.body.insertParagraph(
                `A summary of BGB-16673 drug substance batches manufactured is provided in Table 1. ` +
                `The corresponding batch analysis results for clinical batches are provided in Table 2 and Table 3, respectively.`,
                Word.InsertLocation.end
            );
            introPara.font.name = 'Times New Roman';
            introPara.font.size = 12;
            introPara.font.color = '#000000';
            introPara.font.bold = false;
            introPara.font.italic = false;
            introPara.spaceAfter = 6;
            introPara.spaceBefore = 0;
            introPara.alignment = Word.Alignment.left;

            // 说明段落 - 应用C-Body Text样式
            const notePara = context.document.body.insertParagraph(
                `The acceptance criteria shown were those effective at the time of testing, unless otherwise specified. ` +
                `Individual relative retention times (RRTs) shown for single unspecified impurities that are variable across ` +
                `batches are shown as ranges, if applicable. Any single unspecified impurities below the reporting threshold ` +
                `(0.05%) across all batches are not reported.`,
                Word.InsertLocation.end
            );
            notePara.font.name = 'Times New Roman';
            notePara.font.size = 12;
            notePara.font.color = '#000000';
            notePara.font.bold = false;
            notePara.font.italic = false;
            notePara.spaceAfter = 12;
            notePara.spaceBefore = 0;
            notePara.alignment = Word.Alignment.left;

            // 同步后创建表格
            await context.sync();
            
            // 创建表格1: Overview - 使用实际批次数据
            await createTable1Overview(context);
            await context.sync();
            
            // 创建表格2: Batch Analysis - 使用实际批次数据
            await createTable2BatchAnalysis(context);
            await context.sync();
            
            // 创建表格3: Batch Results - 使用实际批次数据
            await createTable3BatchResults(context);
            await context.sync();
            
        } catch (error) {
            console.error('创建批次分析部分时出错:', error);
            throw error;
        }
    };

    // 创建表格1: Overview of Drug Substance Batches - 应用模板格式
    const createTable1Overview = async (context: Word.RequestContext) => {
        try {
            // 表格标题 - 应用Caption样式
            const table1Caption = context.document.body.insertParagraph('Table 1:	Overview of BGB-16673 Drug Substance Batches', Word.InsertLocation.end);
            table1Caption.font.name = 'Times New Roman';
            table1Caption.font.size = 12;
            table1Caption.font.bold = true;
            table1Caption.font.color = '#000000';
            table1Caption.spaceAfter = 6;
            table1Caption.spaceBefore = 12;
            table1Caption.alignment = Word.Alignment.left;
            table1Caption.font.italic = false;

            // 创建表格 - 实际行数基于批次数据
            const actualBatchCount = batchDataList.length;
            const table1 = context.document.body.insertTable(actualBatchCount + 1, 5, Word.InsertLocation.end);
            table1.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
            
            // 应用表格样式
            table1.font.name = 'Times New Roman';
            table1.font.size = 10;
            table1.font.color = '#000000';

            // 设置表头 - 应用模板表头格式
            const headers = ['Batch Number', 'Batch Size (kg)', 'Date of Manufacture', 'Manufacturer', 'Use(s)'];
            headers.forEach((header, index) => {
                const cell = table1.getCell(0, index);
                cell.value = header;
                cell.body.font.bold = true;
                cell.body.font.name = 'Times New Roman';
                cell.body.font.size = 10;
                cell.body.font.color = '#000000';
                cell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
            });

            // 填充实际批次数据
            batchDataList.forEach((batch, rowIndex) => {
                console.log(`填充表格1第${rowIndex + 1}行数据:`, batch);
                
                const manufacturerShort = batch.manufacturer.includes('Changzhou SynTheAll') ? 'Changzhou STA' : batch.manufacturer;
                const rowData = [
                    batch.batch_number || 'N/A',
                    'TBD', // 批次大小暂时标记为TBD
                    batch.manufacture_date || 'N/A',
                    manufacturerShort,
                    'Clinical batch'
                ];

                rowData.forEach((data, colIndex) => {
                    const cell = table1.getCell(rowIndex + 1, colIndex);
                    cell.value = data;
                    cell.body.font.name = 'Times New Roman';
                    cell.body.font.size = 10;
                    cell.body.font.color = '#000000';
                    cell.body.font.bold = false;
                    cell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
                });
            });

            // 表格后空行 - 应用C-Body Text样式
            const spacePara = context.document.body.insertParagraph('', Word.InsertLocation.end);
            spacePara.spaceAfter = 12;
            spacePara.spaceBefore = 6;
            
            console.log(`表格1创建完成，包含${actualBatchCount}个批次数据`);
        } catch (error) {
            console.error('创建表格1时出错:', error);
            throw error;
        }
    };

    // 创建表格2: Batch Analysis for GMP Batches - 使用实际批次数据
    const createTable2BatchAnalysis = async (context: Word.RequestContext) => {
        try {
            // 表格标题
            const table2Caption = context.document.body.insertParagraph('Table 2: Batch Analysis for GMP Batches of BGB-16673 Drug Substance', Word.InsertLocation.end);
            table2Caption.font.name = 'Times New Roman';
            table2Caption.font.size = 12;
            table2Caption.font.bold = true;
            table2Caption.spaceAfter = 6;
            table2Caption.font.italic = false;

            // 测试参数定义
            const testParameters = [
                { name: "Appearance -- visual inspection", criterion: "Light yellow to yellow powder", key: "Appearance -- visual inspection" },
                { name: "Identification", criterion: "", key: "" },
                { name: "IR", criterion: "Conforms to reference standard", key: "IR" },
                { name: "HPLC", criterion: "Conforms to reference standard", key: "HPLC" },
                { name: "Assay -- HPLC (on anhydrous basis, %w/w)", criterion: "97.0-103.0", key: "Assay -- HPLC (on anhydrous basis, %w/w)" },
                { name: "Organic Impurities -- HPLC (%w/w)", criterion: "", key: "" },
                { name: "Single unspecified impurity", criterion: "≤ 0.50", key: "Single unspecified impurity" },
                { name: "BGB-24860", criterion: "", key: "BGB-24860" },
                { name: "RRT 0.56", criterion: "", key: "RRT 0.56" },
                { name: "RRT 0.70", criterion: "", key: "RRT 0.70" },
                { name: "RRT 0.72-0.73", criterion: "", key: "RRT 0.72-0.73" },
                { name: "RRT 0.76", criterion: "", key: "RRT 0.76" },
                { name: "RRT 0.80", criterion: "", key: "RRT 0.80" },
                { name: "RRT 1.10", criterion: "", key: "RRT 1.10" },
                { name: "Total impurities", criterion: "≤ 2.0", key: "Total impurities" },
                { name: "Enantiomeric Impurity -- HPLC (%w/w)", criterion: "≤ 1.0", key: "Enantiomeric Impurity -- HPLC (%w/w)" },
                { name: "Residual Solvents -- GC (ppm)", criterion: "", key: "" },
                { name: "Dichloromethane", criterion: "≤ 600", key: "Dichloromethane" },
                { name: "Ethyl acetate", criterion: "≤ 5000", key: "Ethyl acetate" },
                { name: "Isopropanol", criterion: "≤ 5000", key: "Isopropanol" },
                { name: "Methanol", criterion: "≤ 3000", key: "Methanol" },
                { name: "Tetrahydrofuran", criterion: "≤ 720", key: "Tetrahydrofuran" }
            ];

            // 创建表格 - 使用实际批次数量
            const actualBatchCount = batchDataList.length;
            const table2 = context.document.body.insertTable(testParameters.length + 1, actualBatchCount + 2, Word.InsertLocation.end);
            table2.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
            table2.font.name = 'Times New Roman';
            table2.font.size = 9;
            table2.font.color = '#000000'; 
            table2.font.bold = false; 
            table2.font.italic = false;
            
            // 设置表头
            const headerCell1 = table2.getCell(0, 0);
            headerCell1.value = 'Test Parameter';
            headerCell1.body.font.bold = true;
            headerCell1.body.paragraphs.getFirst().alignment = Word.Alignment.centered;

            const headerCell2 = table2.getCell(0, 1);
            headerCell2.value = 'Acceptance Criterion';
            headerCell2.body.font.bold = true;
            headerCell2.body.paragraphs.getFirst().alignment = Word.Alignment.centered;

            // 批次号表头 - 使用实际批次数据
            batchDataList.forEach((batch, index) => {
                const headerCell = table2.getCell(0, index + 2);
                headerCell.value = batch.batch_number || `Batch ${index + 1}`;
                headerCell.body.font.bold = true;
                headerCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
            });

            // 填充测试参数数据 - 使用实际测试结果
            testParameters.forEach((param, rowIndex) => {
                // 测试参数名称
                const paramCell = table2.getCell(rowIndex + 1, 0);
                paramCell.value = param.name;
                if (param.name.includes('--')) {
                    paramCell.body.font.bold = true;;
                }
                paramCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered; // 与表格3数据单元格一致

                // 接受标准
                const criterionCell = table2.getCell(rowIndex + 1, 1);
                criterionCell.value = param.criterion;
                criterionCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;

                // 批次结果 - 使用实际提取的数据
                batchDataList.forEach((batch, batchIndex) => {
                    const resultCell = table2.getCell(rowIndex + 1, batchIndex + 2);
                    let result = '';
                    
                    if (param.key && batch.test_results && batch.test_results[param.key]) {
                        result = batch.test_results[param.key];
                        console.log(`表格2: ${param.key} = ${result} (批次: ${batch.batch_number})`);
                    } else if (!param.key || param.criterion === '') {
                        result = ''; // 空白行
                    } else {
                        result = 'TBD'; // 未找到数据时显示TBD
                    }
                    
                    resultCell.value = result;
                    resultCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
                });
            });

            // 表格后空行
            context.document.body.insertParagraph('', Word.InsertLocation.end);

            // 创建继续表格部分
            await createTable2Continued(context);
            
            console.log(`表格2创建完成，包含${actualBatchCount}个批次的${testParameters.length}个测试参数`);
        } catch (error) {
            console.error('创建表格2时出错:', error);
            throw error;
        }
    };

    // 创建表格2继续部分 - 使用实际批次数据
    const createTable2Continued = async (context: Word.RequestContext) => {
        try {
            // 继续表格标题
            const continuedCaption = context.document.body.insertParagraph('Table 2: Batch Analysis for GMP Batches of BGB-16673 Drug Substance (Continued)', Word.InsertLocation.end);
            continuedCaption.font.name = 'Times New Roman';
            continuedCaption.font.size = 12;
            continuedCaption.font.bold = true;
            continuedCaption.spaceAfter = 6;
            continuedCaption.font.italic = false;

            const continuedParameters = [
                { name: "Residue on Ignition (%w/w)", criterion: "≤ 0.2", key: "Residue on Ignition (%w/w)" },
                { name: "Elemental Impurities -- ICP-MS", criterion: "", key: "" },
                { name: "Palladium (ppm)", criterion: "≤ 25", key: "Palladium (ppm)" },
                { name: "Polymorphic Form -- XRPD", criterion: "Conforms to reference standard", key: "Polymorphic Form -- XRPD" },
                { name: "Water Content -- KF (%w/w)", criterion: "Report result", key: "Water Content -- KF (%w/w)" }
            ];

            // 创建继续表格 - 使用实际批次数量
            const actualBatchCount = batchDataList.length;
            const continuedTable = context.document.body.insertTable(continuedParameters.length + 1, actualBatchCount + 2, Word.InsertLocation.end);
            continuedTable.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
            continuedTable.font.name = 'Times New Roman';
            continuedTable.font.size = 9;
            continuedTable.font.color = '#000000'; // 与表格3一致
            continuedTable.font.bold = false; // 与表格3一致
            continuedTable.font.italic = false; // 与表格3一致            

            // 设置表头（重复）
            const headerCell1 = continuedTable.getCell(0, 0);
            headerCell1.value = 'Test Parameter';
            headerCell1.body.font.bold = true;
            headerCell1.body.paragraphs.getFirst().alignment = Word.Alignment.centered;

            const headerCell2 = continuedTable.getCell(0, 1);
            headerCell2.value = 'Acceptance Criterion';
            headerCell2.body.font.bold = true;
            headerCell2.body.paragraphs.getFirst().alignment = Word.Alignment.centered;

            // 批次号表头
            batchDataList.forEach((batch, index) => {
                const headerCell = continuedTable.getCell(0, index + 2);
                headerCell.value = batch.batch_number || `Batch ${index + 1}`;
                headerCell.body.font.bold = true;
                headerCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
            });

            // 填充继续参数数据 - 使用实际测试结果
            continuedParameters.forEach((param, rowIndex) => {
                const paramCell = continuedTable.getCell(rowIndex + 1, 0);
                paramCell.value = param.name;
                if (param.name.includes('--')) {
                    paramCell.body.font.bold = true;
                }
                paramCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;

                const criterionCell = continuedTable.getCell(rowIndex + 1, 1);
                criterionCell.value = param.criterion;
                criterionCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;

                batchDataList.forEach((batch, batchIndex) => {
                    const resultCell = continuedTable.getCell(rowIndex + 1, batchIndex + 2);
                    let result = '';
                    
                    if (param.key && batch.test_results && batch.test_results[param.key]) {
                        result = batch.test_results[param.key];
                        console.log(`表格2继续: ${param.key} = ${result} (批次: ${batch.batch_number})`);
                    } else if (!param.key || param.criterion === '') {
                        result = ''; // 空白行
                    } else {
                        result = 'TBD'; // 未找到数据时显示TBD
                    }
                    
                    resultCell.value = result;
                    resultCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
                });
            });

            // 添加缩写说明 - 应用C-Footnote样式
            const abbreviations = context.document.body.insertParagraph(
                'Abbreviations: GC = gas chromatography; HPLC = high-performance liquid chromatography; ' +
                'ICP-MS = inductively coupled plasma mass spectrometry; IR = infrared spectroscopy; ' +
                'KF = Karl Fischer; ND = not detected; Pd = Palladium; RRT = relative retention time; ' +
                'XRPD = X‑ray powder diffraction.',
                Word.InsertLocation.end
            );
            abbreviations.font.name = 'Times New Roman';
            abbreviations.font.size = 9;
            abbreviations.font.color = '#000000';
            abbreviations.font.italic = true;
            abbreviations.font.bold = false;
            abbreviations.spaceAfter = 12;
            abbreviations.spaceBefore = 6;
            abbreviations.alignment = Word.Alignment.left;

            // 表格后空行 - 应用C-Body Text样式
            const spacePara = context.document.body.insertParagraph('', Word.InsertLocation.end);
            spacePara.spaceAfter = 12;
            spacePara.spaceBefore = 0;
        } catch (error) {
            console.error('创建表格2继续部分时出错:', error);
            throw error;
        }
    };

    // 创建表格3: Batch Results for GMP Batches - 使用实际批次数据
    const createTable3BatchResults = async (context: Word.RequestContext) => {
        try {
            // 表格标题
            const table3Caption = context.document.body.insertParagraph('Table 3: Batch Results for GMP Batches of BGB-16673 Drug Substance', Word.InsertLocation.end);
            table3Caption.font.name = 'Times New Roman';
            table3Caption.font.size = 12;
            table3Caption.font.bold = true;
            table3Caption.spaceAfter = 6;
            table3Caption.font.italic = false;

            // 使用最新批次作为示例，或者如果只有一个批次就使用第一个
            const selectedBatch = batchDataList.length > 0 ? batchDataList[batchDataList.length - 1] : null;
            
            if (!selectedBatch) {
                console.warn('没有批次数据可用于表格3');
                return;
            }

            const table3Parameters = [
                { name: "Appearance -- visual inspection", criterion: "Light yellow to yellow powder", key: "Appearance -- visual inspection" },
                { name: "Identification", criterion: "", key: "" },
                { name: "IR", criterion: "Conforms to reference standard", key: "IR" },
                { name: "HPLC", criterion: "Conforms to reference standard", key: "HPLC" },
                { name: "Assay -- HPLC (on anhydrous basis, % w/w)", criterion: "97.0-103.0", key: "Assay -- HPLC (on anhydrous basis, %w/w)" },
                { name: "Organic Impurities -- HPLC (% w/w)", criterion: "", key: "" },
                { name: "Single unspecified impurity", criterion: "≤ 0.30", key: "Single unspecified impurity" },
                { name: "RRT 0.83", criterion: "", key: "RRT 0.83" },
                { name: "Total impurities", criterion: "≤ 2.0", key: "Total impurities" },
                { name: "Enantiomeric Impurity -- HPLC (% w/w)", criterion: "≤ 0.5", key: "Enantiomeric Impurity -- HPLC (%w/w)" },
                { name: "Residual Solvents -- GC (ppm)", criterion: "", key: "" },
                { name: "Dichloromethane", criterion: "≤ 600", key: "Dichloromethane" },
                { name: "Ethyl acetate", criterion: "≤ 5000", key: "Ethyl acetate" },
                { name: "Isopropanol", criterion: "≤ 5000", key: "Isopropanol" },
                { name: "Methanol", criterion: "≤ 3000", key: "Methanol" },
                { name: "Tetrahydrofuran", criterion: "≤ 720", key: "Tetrahydrofuran" },
                { name: "Inorganic Impurities", criterion: "", key: "" },
                { name: "Residue on ignition (% w/w)", criterion: "≤ 0.2", key: "Residue on Ignition (%w/w)" },
                { name: "Elemental impurities -- ICP-MS (Pd) (ppm)", criterion: "≤ 25", key: "Palladium (ppm)" },
                { name: "Polymorphic Form -- XRPD", criterion: "Conforms to reference standard", key: "Polymorphic Form -- XRPD" },
                { name: "Water Content -- KF (% w/w)", criterion: "≤ 3.5", key: "Water Content -- KF (%w/w)" }
            ];

            // 创建表格3
            const table3 = context.document.body.insertTable(table3Parameters.length + 1, 3, Word.InsertLocation.end);
            table3.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
            table3.font.name = 'Times New Roman';
            table3.font.size = 9;
            table3.font.color = '#000000';
            table3.font.bold = false;
            table3.font.italic = false;

            // 设置表头
            const headers3 = ['Test Parameter', 'Acceptance Criterion', selectedBatch.batch_number || 'Latest Batch'];
            headers3.forEach((header, index) => {
                const cell = table3.getCell(0, index);
                cell.value = header;
                cell.body.font.bold = true;
                cell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
            });

            // 填充表格3数据 - 使用实际测试结果
            table3Parameters.forEach((param, rowIndex) => {
                const paramCell = table3.getCell(rowIndex + 1, 0);
                paramCell.value = param.name;
                if (param.name.includes('--')) {
                    paramCell.body.font.bold = true;
                }

                const criterionCell = table3.getCell(rowIndex + 1, 1);
                criterionCell.value = param.criterion;

                const resultCell = table3.getCell(rowIndex + 1, 2);
                let result = '';
                
                if (param.key && selectedBatch.test_results && selectedBatch.test_results[param.key]) {
                    result = selectedBatch.test_results[param.key];
                    console.log(`表格3: ${param.key} = ${result} (批次: ${selectedBatch.batch_number})`);
                } else if (!param.key || param.criterion === '') {
                    result = ''; // 空白行
                } else {
                    result = 'TBD'; // 未找到数据时显示TBD
                }
                
                resultCell.value = result;
                resultCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
            });

            // 添加缩写说明 - 应用C-Footnote样式
            const abbreviations3 = context.document.body.insertParagraph(
                'Abbreviations: GC = gas chromatography; HPLC = high-performance liquid chromatography; ' +
                'ICP-MS = inductively coupled plasma mass spectrometry; IR = infrared spectroscopy; ' +
                'KF = Karl Fischer; ND = not detected; Pd = Palladium; RRT = relative retention time; ' +
                'XRPD = X‑ray powder diffraction.',
                Word.InsertLocation.end
            );
            abbreviations3.font.name = 'Times New Roman';
            abbreviations3.font.size = 9;
            abbreviations3.font.color = '#000000';
            abbreviations3.font.italic = true;
            abbreviations3.font.bold = false;
            abbreviations3.spaceAfter = 12;
            abbreviations3.spaceBefore = 6;
            abbreviations3.alignment = Word.Alignment.left;
            
            console.log(`表格3创建完成，使用批次: ${selectedBatch.batch_number}`);
        } catch (error) {
            console.error('创建表格3时出错:', error);
            throw error;
        }
    };

    // 简化版插入功能（保持向后兼容）
    const handleInsertToWord = async () => {
        if (!state.selectedTemplate || batchDataList.length === 0) {
            setState(prev => ({ 
                ...prev, 
                error: 'No batch analysis data available to insert.' 
            }));
            return;
        }

        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                
                // 生成基于模板的批次分析表格
                const tableHtml = generateBatchAnalysisTable(batchDataList);
                
                selection.insertHtml(tableHtml, Word.InsertLocation.replace);
                await context.sync();
            });
            
            setState(prev => ({ 
                ...prev, 
                error: undefined 
            }));

            setProcessingStatus('AIMTA Batch Analysis Tables inserted successfully!');
            
        } catch (error) {
            setState(prev => ({ 
                ...prev, 
                error: 'Failed to insert data into Word document.' 
            }));
        } finally {
            setState(prev => ({ ...prev, isLoading: false }));
        }
    };

    const generateBatchAnalysisTable = (batchData: BatchData[]): string => {
        if (!batchData.length) return '';

        // Table 1: Overview of Drug Substance Batches
        let table1Html = `
        <h2>Table 1: Overview of BGB-16673 Drug Substance Batches</h2>
        <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%;">
            <thead>
                <tr style="background-color: #D9E1F2; font-weight: bold;">
                    <td style="padding: 8px; border: 1px solid #000;">Batch Number</td>
                    <td style="padding: 8px; border: 1px solid #000;">Batch Size (kg)</td>
                    <td style="padding: 8px; border: 1px solid #000;">Date of Manufacture</td>
                    <td style="padding: 8px; border: 1px solid #000;">Manufacturer</td>
                    <td style="padding: 8px; border: 1px solid #000;">Use(s)</td>
                </tr>
            </thead>
            <tbody>`;

        batchData.forEach(batch => {
            const manufacturerShort = batch.manufacturer.includes('Changzhou SynTheAll') ? 'Changzhou STA' : batch.manufacturer;
            table1Html += `
                <tr>
                    <td style="padding: 8px; border: 1px solid #000;">${batch.batch_number}</td>
                    <td style="padding: 8px; border: 1px solid #000;">TBD</td>
                    <td style="padding: 8px; border: 1px solid #000;">${batch.manufacture_date}</td>
                    <td style="padding: 8px; border: 1px solid #000;">${manufacturerShort}</td>
                    <td style="padding: 8px; border: 1px solid #000;">Clinical batch</td>
                </tr>`;
        });

        table1Html += `
            </tbody>
        </table>
        <p></p>`;

        // Table 2: Batch Analysis for GMP Batches
        let table2Html = `
        <h2>Table 2: Batch Analysis for GMP Batches of BGB-16673 Drug Substance</h2>
        <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%;">
            <thead>
                <tr style="background-color: #D9E1F2; font-weight: bold;">
                    <td style="padding: 8px; border: 1px solid #000; width: 25%;">Test Parameter</td>
                    <td style="padding: 8px; border: 1px solid #000; width: 20%;">Acceptance Criterion</td>`;

        // Add batch number headers
        batchData.forEach(batch => {
            table2Html += `<td style="padding: 8px; border: 1px solid #000; text-align: center;">${batch.batch_number}</td>`;
        });

        table2Html += `
                </tr>
            </thead>
            <tbody>`;

        // Test parameters from template - Table 2
        const testParameters = [
            { name: "Appearance -- visual inspection", criterion: "Light yellow to yellow powder", key: "Appearance -- visual inspection" },
            { name: "Identification", criterion: "", key: "" },
            { name: "IR", criterion: "Conforms to reference standard", key: "IR" },
            { name: "HPLC", criterion: "Conforms to reference standard", key: "HPLC" },
            { name: "Assay -- HPLC (on anhydrous basis, %w/w)", criterion: "97.0-103.0", key: "Assay -- HPLC (on anhydrous basis, %w/w)" },
            { name: "Organic Impurities -- HPLC (%w/w)", criterion: "", key: "" },
            { name: "Single unspecified impurity", criterion: "≤ 0.50", key: "Single unspecified impurity" },
            { name: "BGB-24860", criterion: "", key: "BGB-24860" },
            { name: "RRT 0.56", criterion: "", key: "RRT 0.56" },
            { name: "RRT 0.70", criterion: "", key: "RRT 0.70" },
            { name: "RRT 0.72-0.73", criterion: "", key: "RRT 0.72-0.73" },
            { name: "RRT 0.76", criterion: "", key: "RRT 0.76" },
            { name: "RRT 0.80", criterion: "", key: "RRT 0.80" },
            { name: "RRT 1.10", criterion: "", key: "RRT 1.10" },
            { name: "Total impurities", criterion: "≤ 2.0", key: "Total impurities" },
            { name: "Enantiomeric Impurity -- HPLC (%w/w)", criterion: "≤ 1.0", key: "Enantiomeric Impurity -- HPLC (%w/w)" },
            { name: "Residual Solvents -- GC (ppm)", criterion: "", key: "" },
            { name: "Dichloromethane", criterion: "≤ 600", key: "Dichloromethane" },
            { name: "Ethyl acetate", criterion: "≤ 5000", key: "Ethyl acetate" },
            { name: "Isopropanol", criterion: "≤ 5000", key: "Isopropanol" },
            { name: "Methanol", criterion: "≤ 3000", key: "Methanol" },
            { name: "Tetrahydrofuran", criterion: "≤ 720", key: "Tetrahydrofuran" }
        ];

        testParameters.forEach(param => {
            table2Html += `
                <tr>
                    <td style="padding: 8px; border: 1px solid #000; font-weight: bold;">${param.name}</td>
                    <td style="padding: 8px; border: 1px solid #000;">${param.criterion}</td>`;

            batchData.forEach(batch => {
                const result = param.key ? (batch.test_results[param.key] || 'TBD') : '';
                table2Html += `<td style="padding: 8px; border: 1px solid #000;">${result}</td>`;
            });

            table2Html += `</tr>`;
        });

        table2Html += `
            </tbody>
        </table>
        <p></p>
        
        <h3>Table 2: Batch Analysis for GMP Batches of BGB-16673 Drug Substance (Continued)</h3>
        <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%;">
            <thead>
                <tr style="background-color: #D9E1F2; font-weight: bold;">
                    <td style="padding: 8px; border: 1px solid #000; width: 25%;">Test Parameter</td>
                    <td style="padding: 8px; border: 1px solid #000; width: 20%;">Acceptance Criterion</td>`;

        // Add batch number headers for continued table
        batchData.forEach(batch => {
            table2Html += `<td style="padding: 8px; border: 1px solid #000; text-align: center;">${batch.batch_number}</td>`;
        });

        table2Html += `
                </tr>
            </thead>
            <tbody>`;

        // Table 2 Continued parameters
        const continuedParameters = [
            { name: "Residue on Ignition (%w/w)", criterion: "≤ 0.2", key: "Residue on Ignition (%w/w)" },
            { name: "Elemental Impurities -- ICP-MS", criterion: "", key: "" },
            { name: "Palladium (ppm)", criterion: "≤ 25", key: "Palladium (ppm)" },
            { name: "Polymorphic Form -- XRPD", criterion: "Conforms to reference standard", key: "Polymorphic Form -- XRPD" },
            { name: "Water Content -- KF (%w/w)", criterion: "Report result", key: "Water Content -- KF (%w/w)" }
        ];

        continuedParameters.forEach(param => {
            table2Html += `
                <tr>
                    <td style="padding: 8px; border: 1px solid #000; font-weight: bold;">${param.name}</td>
                    <td style="padding: 8px; border: 1px solid #000;">${param.criterion}</td>`;

            batchData.forEach(batch => {
                const result = param.key ? (batch.test_results[param.key] || 'TBD') : '';
                table2Html += `<td style="padding: 8px; border: 1px solid #000;">${result}</td>`;
            });

            table2Html += `</tr>`;
        });

        table2Html += `
            </tbody>
        </table>
        <p></p>`;

        // Add abbreviations for Table 2
        table2Html += `
        <p style="font-size: 9pt; font-style: italic;">
        <strong>Abbreviations:</strong> GC = gas chromatography; HPLC = high-performance liquid chromatography; 
        ICP-MS = inductively coupled plasma mass spectrometry; IR = infrared spectroscopy; KF = Karl Fischer; 
        ND = not detected; Pd = Palladium; RRT = relative retention time; XRPD = X‑ray powder diffraction.
        </p>
        <p></p>`;

        // Table 3: Batch Results for GMP Batches (Single batch format from template)
        let table3Html = `
        <h2>Table 3: Batch Results for GMP Batches of BGB-16673 Drug Substance</h2>`;

        // Generate Table 3 for the latest/most recent batch (or you can modify this logic)
        const latestBatch = batchData[batchData.length - 1]; // Take the last batch as example

        table3Html += `
        <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%;">
            <thead>
                <tr style="background-color: #D9E1F2; font-weight: bold;">
                    <td style="padding: 8px; border: 1px solid #000; width: 40%;">Test Parameter</td>
                    <td style="padding: 8px; border: 1px solid #000; width: 30%;">Acceptance Criterion</td>
                    <td style="padding: 8px; border: 1px solid #000; width: 30%;">${latestBatch.batch_number}</td>
                </tr>
            </thead>
            <tbody>`;

        // Table 3 parameters (based on template structure)
        const table3Parameters = [
            { name: "Appearance -- visual inspection", criterion: "Light yellow to yellow powder", key: "Appearance -- visual inspection" },
            { name: "Identification", criterion: "", key: "" },
            { name: "IR", criterion: "Conforms to reference standard", key: "IR" },
            { name: "HPLC", criterion: "Conforms to reference standard", key: "HPLC" },
            { name: "Assay -- HPLC (on anhydrous basis, % w/w)", criterion: "97.0-103.0", key: "Assay -- HPLC (on anhydrous basis, %w/w)" },
            { name: "Organic Impurities -- HPLC (% w/w)", criterion: "", key: "" },
            { name: "Single unspecified impurity", criterion: "≤ 0.30", key: "Single unspecified impurity" },
            { name: "RRT 0.83", criterion: "", key: "RRT 0.83" },
            { name: "Total impurities", criterion: "≤ 2.0", key: "Total impurities" },
            { name: "Enantiomeric Impurity -- HPLC (% w/w)", criterion: "≤ 0.5", key: "Enantiomeric Impurity -- HPLC (%w/w)" },
            { name: "Residual Solvents -- GC (ppm)", criterion: "", key: "" },
            { name: "Dichloromethane", criterion: "≤ 600", key: "Dichloromethane" },
            { name: "Ethyl acetate", criterion: "≤ 5000", key: "Ethyl acetate" },
            { name: "Isopropanol", criterion: "≤ 5000", key: "Isopropanol" },
            { name: "Methanol", criterion: "≤ 3000", key: "Methanol" },
            { name: "Tetrahydrofuran", criterion: "≤ 720", key: "Tetrahydrofuran" },
            { name: "Inorganic Impurities", criterion: "", key: "" },
            { name: "Residue on ignition (% w/w)", criterion: "≤ 0.2", key: "Residue on Ignition (%w/w)" },
            { name: "Elemental impurities -- ICP-MS (Pd) (ppm)", criterion: "≤ 25", key: "Palladium (ppm)" },
            { name: "Polymorphic Form -- XRPD", criterion: "Conforms to reference standard", key: "Polymorphic Form -- XRPD" },
            { name: "Water Content -- KF (% w/w)", criterion: "≤ 3.5", key: "Water Content -- KF (%w/w)" }
        ];

        table3Parameters.forEach(param => {
            const result = param.key ? (latestBatch.test_results[param.key] || 'TBD') : '';
            table3Html += `
                <tr>
                    <td style="padding: 8px; border: 1px solid #000; font-weight: bold;">${param.name}</td>
                    <td style="padding: 8px; border: 1px solid #000;">${param.criterion}</td>
                    <td style="padding: 8px; border: 1px solid #000;">${result}</td>
                </tr>`;
        });

        table3Html += `
            </tbody>
        </table>
        <p></p>`;

        // Add abbreviations for Table 3
        const finalAbbreviations = `
        <p style="font-size: 9pt; font-style: italic;">
        <strong>Abbreviations:</strong> GC = gas chromatography; HPLC = high-performance liquid chromatography; 
        ICP-MS = inductively coupled plasma mass spectrometry; IR = infrared spectroscopy; KF = Karl Fischer; 
        ND = not detected; Pd = Palladium; RRT = relative retention time; XRPD = X‑ray powder diffraction.
        </p>`;

        return table1Html + table2Html + table3Html + finalAbbreviations;
    };

    return (
        <Stack className="app-container" tokens={{ childrenGap: 20 }}>
            <Stack.Item>
                <Text variant="xLarge" className="app-title">AIMTA Batch Analysis Processor</Text>
            </Stack.Item>

                        {/* 连接状态显示 - 新增 */}
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

            {/* 如果连接失败，显示重试按钮 - 新增 */}
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


            {state.error && (
                <MessageBar 
                    messageBarType={MessageBarType.error}
                    onDismiss={() => setState(prev => ({ ...prev, error: undefined }))}
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
                                
                                {/* 数据管理按钮组 */}
                                <Stack horizontal tokens={{ childrenGap: 10 }}>
                                    <DefaultButton
                                        text="Force Reprocess"
                                        iconProps={{ iconName: 'Refresh' }}
                                        onClick={handleForceReprocess}
                                        disabled={state.isLoading}
                                        styles={{ 
                                            root: { 
                                                flex: 1,
                                                backgroundColor: '#FF8C00',
                                                color: 'white',
                                                borderColor: '#FF8C00'
                                            } 
                                        }}
                                    />
                                    <DefaultButton
                                        text="Clear Cache"
                                        iconProps={{ iconName: 'Delete' }}
                                        onClick={handleClearCache}
                                        disabled={state.isLoading}
                                        styles={{ 
                                            root: { 
                                                flex: 1,
                                                backgroundColor: '#DC143C',
                                                color: 'white',
                                                borderColor: '#DC143C'
                                            } 
                                        }}
                                    />
                                </Stack>
                                
                                {/* 保留原有的简单插入功能 */}
                                <DefaultButton
                                    text="Insert Tables Only (Legacy)"
                                    iconProps={{ iconName: 'Table' }}
                                    onClick={handleInsertToWord}
                                    disabled={state.isLoading}
                                    styles={{ root: { width: '100%' } }}
                                />
                            </Stack>
                        )}
                    </Stack>
                )}
            </Stack>

            <Stack className="help-text">
                <Text variant="small">  - Select compound and template region</Text>
                <Text variant="small">  - Click "Load/Process AIMTA Files" to check database first, then process if needed</Text>
                <Text variant="small">  - "Create Complete AIMTA Document" generates a full formatted document with:</Text>
                <Text variant="small">  - Template-based formatting (Calibri font, proper spacing, colors)</Text>
                <Text variant="small">  - Professional table of contents and list of tables</Text>
                <Text variant="small">  - Headers/footers (with fallback instructions if not supported)</Text>
                <Text variant="small">  - Properly formatted tables with extracted batch data</Text>
                <Text variant="small">  - Data Management:</Text>
                <Text variant="small">  - "Force Reprocess": Re-scan PDFs and update database (orange button)</Text>
                <Text variant="small">  - "Clear Cache": Remove cached data from database (red button)</Text>
                <Text variant="small">  - "Insert Tables Only" provides legacy functionality for simple table insertion</Text>
                <Text variant="small">  - Database caching prevents duplicate processing and improves performance</Text>
            </Stack>
        </Stack>
    );
};