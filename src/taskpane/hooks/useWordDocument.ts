import { useState, useCallback, useRef } from 'react';
import { Template, BatchData } from '../../types';

// 批量操作管理器
class WordBatchOperations {
    private operations: Array<(context: Word.RequestContext) => Promise<void>> = [];
    
    add(operation: (context: Word.RequestContext) => Promise<void>) {
        this.operations.push(operation);
    }
    
    async execute(context: Word.RequestContext) {
        for (const operation of this.operations) {
            await operation(context);
        }
        this.operations = [];
    }
    
    get length() {
        return this.operations.length;
    }
}

export const useWordDocument = () => {
    const [isCreating, setIsCreating] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const batchOps = useRef(new WordBatchOperations());

    // 批量创建段落
    const createParagraphBatch = useCallback((
        texts: string[],
        location: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End",
        formatter?: (paragraph: Word.Paragraph) => void
    ) => {
        batchOps.current.add(async (context) => {
            const body = context.document.body;
            
            for (const text of texts) {
                const paragraph = body.insertParagraph(text, location as Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End");
                if (formatter) {
                    formatter(paragraph);
                }
            }
        });
    }, []);

    // 批量创建表格
    const createTableBatch = useCallback((
        tableData: {
            rows: number;
            cols: number;
            data: string[][];
            formatter?: (table: Word.Table) => void;
        },
        location: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End"
    ) => {
        batchOps.current.add(async (context) => {
            const body = context.document.body;
            const table = body.insertTable(tableData.rows, tableData.cols, location);
            
            // 设置表格数据
            for (let i = 0; i < tableData.data.length; i++) {
                for (let j = 0; j < tableData.data[i].length; j++) {
                    const cell = table.getCell(i, j);
                    cell.value = tableData.data[i][j];
                }
            }
            
            if (tableData.formatter) {
                tableData.formatter(table);
            }
        });
    }, []);

    // 创建完整文档
    const createDocument = useCallback(async (template: Template, batchDataList: BatchData[]) => {
        try {
            setIsCreating(true);
            setError(null);

            await Word.run(async (context) => {
                // 清空文档
                context.document.body.clear();
                
                // 添加目录页操作
                createParagraphBatch(
                    ['Table of Contents'],
                    Word.InsertLocation.end,
                    (p) => {
                        p.alignment = Word.Alignment.centered;
                        p.font.name = 'Times New Roman';
                        p.font.size = 14;
                        p.font.bold = true;
                        p.spaceAfter = 12;
                    }
                );

                // 添加目录条目
                const tocEntries = [
                    'TABLE OF CONTENTS\t1',
                    'LIST OF TABLES\t1',
                    'S.4.4  BATCH ANALYSES\t2'
                ];
                
                createParagraphBatch(
                    tocEntries,
                    Word.InsertLocation.end,
                    (p) => {
                        p.alignment = Word.Alignment.left;
                        p.font.name = 'Times New Roman';
                        p.font.size = 12;
                        p.font.color = '#0000FF';
                    }
                );

                // 插入分页符
                batchOps.current.add(async (ctx) => {
                    ctx.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
                });

                // 创建表格1数据
                const table1Data = [
                    ['Batch Number', 'Batch Size (kg)', 'Date of Manufacture', 'Manufacturer', 'Use(s)'],
                    ...batchDataList.map(batch => [
                        batch.batch_number || 'N/A',
                        'TBD',
                        batch.manufacture_date || 'N/A',
                        batch.manufacturer.includes('Changzhou SynTheAll') ? 'Changzhou STA' : batch.manufacturer,
                        'Clinical batch'
                    ])
                ];

                createTableBatch(
                    {
                        rows: table1Data.length,
                        cols: 5,
                        data: table1Data,
                        formatter: (table) => {
                            table.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
                            table.font.name = 'Times New Roman';
                            table.font.size = 10;
                        }
                    },
                    Word.InsertLocation.end
                );

                // 执行所有批量操作
                await batchOps.current.execute(context);
                await context.sync();
            });

        } catch (error) {
            console.error('Error creating document:', error);
            setError(error instanceof Error ? error.message : 'Failed to create document');
            throw error;
        } finally {
            setIsCreating(false);
        }
    }, [createParagraphBatch, createTableBatch]);

    // 插入表格
    const insertTables = useCallback(async (batchDataList: BatchData[]) => {
        try {
            setIsCreating(true);
            setError(null);

            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                
                // 生成HTML表格
                const tableHtml = generateBatchAnalysisTableHTML(batchDataList);
                
                // 一次性插入所有表格
                selection.insertHtml(tableHtml, Word.InsertLocation.replace);
                
                await context.sync();
            });

        } catch (error) {
            console.error('Error inserting tables:', error);
            setError(error instanceof Error ? error.message : 'Failed to insert tables');
            throw error;
        } finally {
            setIsCreating(false);
        }
    }, []);

    return {
        createDocument,
        insertTables,
        isCreating,
        error
    };
};

// 辅助函数：生成批量表格HTML
function generateBatchAnalysisTableHTML(batchDataList: BatchData[]): string {
    if (!batchDataList.length) return '';

    // 生成所有表格的HTML
    let html = '';
    
    // Table 1
    html += '<h2>Table 1: Overview of BGB-16673 Drug Substance Batches</h2>';
    html += '<table border="1" style="border-collapse: collapse; width: 100%;">';
    html += '<thead><tr style="background-color: #D9E1F2;">';
    html += '<th>Batch Number</th><th>Batch Size (kg)</th><th>Date of Manufacture</th><th>Manufacturer</th><th>Use(s)</th>';
    html += '</tr></thead><tbody>';
    
    batchDataList.forEach(batch => {
        html += '<tr>';
        html += `<td>${batch.batch_number}</td>`;
        html += '<td>TBD</td>';
        html += `<td>${batch.manufacture_date}</td>`;
        html += `<td>${batch.manufacturer.includes('Changzhou SynTheAll') ? 'Changzhou STA' : batch.manufacturer}</td>`;
        html += '<td>Clinical batch</td>';
        html += '</tr>';
    });
    
    html += '</tbody></table><br/>';
    
    return html;
}