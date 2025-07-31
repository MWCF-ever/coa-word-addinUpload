// import React, { useState, useCallback, useMemo } from 'react';
// import { 
//     Stack, Text, TextField, PrimaryButton, DefaultButton,
//     ProgressIndicator, Icon, Separator, MessageBar, MessageBarType
// } from '@fluentui/react';
// import { ExtractedField, API_BASE_URL } from '../../types';
// import axios from 'axios';


// const FIELD_DISPLAY_NAMES: Record<string, string> = {
//     'lot_number': 'Lot Number / 批号',
//     'manufacturer': 'Manufacturer / 生产商',
//     'storage_condition': 'Storage Condition / 储存条件',
//     'Manufacture_Date': 'Manufacture Date / 生产日期'
// };

// const getFieldDisplayName = (fieldName: string): string => {
//     return FIELD_DISPLAY_NAMES[fieldName] || fieldName;
// };

// const getConfidenceColor = (score: number): string => {
//     if (score >= 0.9) return '#107C10'; // Green
//     if (score >= 0.7) return '#FFA500'; // Orange
//     return '#D13438'; // Red
// };



// /**
//  * 从后端API获取用于Word表格的数据
//  */
// const fetchWordTableData = async (documentId: string) => {
//     const apiUrl = `${API_BASE_URL}/documents/${documentId}/word-table-data`;
//     try {
//         const response = await axios.get(apiUrl, { timeout: 30000 });
//         if (!response.data.success) {
//             throw new Error(response.data.error || 'API returned failure');
//         }
//         const tableData = response.data.data.tableData;
//         if (!tableData || !Array.isArray(tableData)) {
//             throw new Error('Invalid table data format received from API');
//         }
//         return tableData;
//     } catch (error) {
//         console.error('Failed to fetch table data:', error);
//         // 重新抛出错误，以便上层函数捕获并处理UI
//         throw error;
//     }
// };

// /**
//  * 根据数据构建HTML表格字符串
//  */
// const buildHtmlTable = (tableData: any[]): string => {
//     const tableRows = tableData.map(item => {
//         let cellValue = item.value || '';
//         if (item.confidence !== undefined && item.confidence > 0) {
//             cellValue += ` (${(item.confidence * 100).toFixed(0)}%)`;
//         }

//         let bgColor = '';
//         if (item.confidence !== undefined) {
//             if (item.confidence < 0.6) bgColor = 'background-color: #FFE6E6;'; // Light Red
//             else if (item.confidence < 0.8) bgColor = 'background-color: #FFF4E6;'; // Light Orange
//         }

//         return `
//             <tr>
//                 <td style="padding: 8px; border: 1px solid #000; font-weight: bold;">${item.field || ''}</td>
//                 <td style="padding: 8px; border: 1px solid #000; ${bgColor}">${cellValue}</td>
//             </tr>
//         `;
//     }).join('');

//     return `
//         <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%;">
//             <thead>
//                 <tr style="background-color: #D9E1F2; font-weight: bold;">
//                     <td style="padding: 8px; border: 1px solid #000; width: 40%;">Field / 字段</td>
//                     <td style="padding: 8px; border: 1px solid #000; width: 60%;">Value / 值</td>
//                 </tr>
//             </thead>
//             <tbody>
//                 ${tableRows}
//             </tbody>
//         </table>
//         <p></p>
//     `;
// };

// /**
//  * 将HTML内容插入到Word文档的当前光标位置
//  */
// const insertHtmlIntoWord = async (html: string): Promise<void> => {
//     return Word.run(async (context) => {
//         try {
//             const range = context.document.getSelection();
//             range.insertHtml(html, Word.InsertLocation.after);
//             await context.sync();
//         } catch (wordError: any) {
//             console.error('Word API error:', wordError);
//             throw new Error(`Word operation failed: ${wordError?.message || wordError}`);
//         }
//     });
// };

// /**
//  * 格式化错误消息以显示给用户
//  */
// const formatErrorMessage = (error: any): string => {
//     if (axios.isAxiosError(error)) {
//         if (error.response) {
//             if (error.response.status === 404) return '文档未找到，请重新上传并处理文档。';
//             if (error.response.status === 500) return `服务器错误: ${error.response.data?.detail || '内部错误'}`;
//             return `网络错误 (${error.response.status}): ${error.message}`;
//         }
//         if (error.code === 'ECONNABORTED') {
//             return '请求超时，请检查网络连接后重试。';
//         }
//         return `网络请求失败: ${error.message}`;
//     }
//     if (error instanceof Error) {
//         return error.message;
//     }
//     return '发生未知错误，插入失败。';
// };


// // --- 主组件定义 ---

// interface ResultDisplayProps {
//     extractedData: ExtractedField[];
//     onDataUpdate: (updatedData: ExtractedField[]) => void;
//     documentId?: string;
// }

// // 使用 React.memo 提高性能 ---
// // 如果props没有变化，可以跳过组件的重新渲染。
// export const ResultDisplay: React.FC<ResultDisplayProps> = React.memo(({
//     extractedData,
//     onDataUpdate,
//     documentId
// }) => {
//     const [editedData, setEditedData] = useState<Record<string, string>>({});
    
//     // --- 优化 4: 合并关联的状态 ---
//     const [insertionStatus, setInsertionStatus] = useState({
//         isLoading: false,
//         message: null as string | null,
//         error: null as string | null,
//     });

//     const handleFieldChange = useCallback((fieldName: string, value: string) => {
//         setEditedData(prev => ({
//             ...prev,
//             [fieldName]: value
//         }));
//     }, []);

//     const handleSaveChanges = useCallback(() => {
//         const updatedData = extractedData.map(field => ({
//             ...field,
//             fieldValue: editedData[field.fieldName] ?? field.fieldValue 
//         }));
//         onDataUpdate(updatedData);
//         setEditedData({});
//     }, [extractedData, editedData, onDataUpdate]);

//     const handleInsertToWord = useCallback(async () => {
//         if (!documentId) {
//             setInsertionStatus({ isLoading: false, message: null, error: 'Document ID not available.' });
//             return;
//         }

//         setInsertionStatus({ isLoading: true, message: null, error: null });

//         try {
//             // 步骤 1: 获取数据
//             const tableData = await fetchWordTableData(documentId);
            
//             // 步骤 2: 构建HTML
//             const htmlTable = buildHtmlTable(tableData);
            
//             // 步骤 3: 插入到Word
//             await insertHtmlIntoWord(htmlTable);

//             setInsertionStatus({
//                 isLoading: false,
//                 message: 'COA表格已成功插入到Word文档的当前光标位置！',
//                 error: null
//             });

//         } catch (error: any) {
//             console.error('Failed to insert table into Word:', error);
//             setInsertionStatus({
//                 isLoading: false,
//                 message: null,
//                 error: formatErrorMessage(error)
//             });
//         }
//     }, [documentId]);

//     // 只有在 `editedData` 变化时才重新计算，避免不必要的计算。
//     const hasUnsavedChanges = useMemo(() => Object.keys(editedData).length > 0, [editedData]);

//     return (
//         <Stack tokens={{ childrenGap: 15 }}>
//             <Separator />
//             <Text variant="large" className="section-title">Extracted Data</Text>
            
//             <MessageBar messageBarType={MessageBarType.info}>
//                 Document ID: {documentId || 'Not available'}
//             </MessageBar>
            
//             {insertionStatus.message && (
//                 <MessageBar 
//                     messageBarType={MessageBarType.success}
//                     onDismiss={() => setInsertionStatus(prev => ({ ...prev, message: null }))}
//                 >
//                     {insertionStatus.message}
//                 </MessageBar>
//             )}
            
//             {insertionStatus.error && (
//                 <MessageBar 
//                     messageBarType={MessageBarType.error}
//                     onDismiss={() => setInsertionStatus(prev => ({ ...prev, error: null }))}
//                 >
//                     {insertionStatus.error}
//                 </MessageBar>
//             )}
            
//             {insertionStatus.isLoading && (
//                 <ProgressIndicator label="正在插入表格到Word文档..." />
//             )}
            
//             <Stack tokens={{ childrenGap: 15 }}>
//                 {extractedData.map((field) => (
//                     <Stack key={field.id} tokens={{ childrenGap: 5 }}>
//                         <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
//                             <Text variant="medium">{getFieldDisplayName(field.fieldName)}</Text>
//                             <Stack horizontal tokens={{ childrenGap: 5 }} verticalAlign="center">
//                                 <Icon 
//                                     iconName="Info" 
//                                     style={{ 
//                                         color: getConfidenceColor(field.confidenceScore),
//                                         fontSize: 12
//                                     }} 
//                                 />
//                                 <Text 
//                                     variant="small" 
//                                     style={{ color: getConfidenceColor(field.confidenceScore) }}
//                                 >
//                                     {(field.confidenceScore * 100).toFixed(0)}% confidence
//                                 </Text>
//                             </Stack>
//                         </Stack>
                        
//                         <TextField
//                             value={editedData[field.fieldName] ?? field.fieldValue}
//                             onChange={(e, newValue) => handleFieldChange(field.fieldName, newValue || '')}
//                             multiline={field.fieldName === 'storage_condition'}
//                             rows={field.fieldName === 'storage_condition' ? 3 : 1}
//                             styles={{
//                                 fieldGroup: {
//                                     backgroundColor: field.confidenceScore < 0.7 ? '#FFF4E6' : undefined
//                                 }
//                             }}
//                         />
                        
//                         {field.originalText && (
//                             <Text variant="small" style={{ color: '#605E5C', fontStyle: 'italic' }}>
//                                 Original: "{field.originalText}"
//                             </Text>
//                         )}
//                     </Stack>
//                 ))}
//             </Stack>

//             <Separator />

//             <Stack horizontal tokens={{ childrenGap: 10 }}>
//                 <PrimaryButton
//                     text={insertionStatus.isLoading ? "正在插入..." : "插入表格到Word"}
//                     iconProps={{ iconName: 'Table' }}
//                     onClick={handleInsertToWord}
//                     disabled={insertionStatus.isLoading || !documentId}
//                     styles={{ root: { flex: 1 } }}
//                 />
//                 <DefaultButton
//                     text="保存更改"
//                     iconProps={{ iconName: 'Save' }}
//                     onClick={handleSaveChanges}
//                     disabled={!hasUnsavedChanges || insertionStatus.isLoading}
//                     styles={{ root: { flex: 1 } }}
//                 />
//             </Stack>

//             {/* Help text section remains the same */}
//             <Stack className="help-text">
//                 <Text variant="small">• 表格将插入到Word文档的当前光标位置</Text>
//                 <Text variant="small">• 请确保在Word文档中选择合适的插入位置</Text>
//                 <Text variant="small">• 低置信度字段用黄色/红色背景标记，请仔细检查</Text>
//                 <Text variant="small">• 点击"保存更改"以在插入前应用修改</Text>
//             </Stack>
//         </Stack>
//     );
// });