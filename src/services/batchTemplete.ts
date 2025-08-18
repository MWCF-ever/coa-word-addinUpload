import { BatchData } from '../types/';

// ========== 样式配置 ==========
interface ParagraphStyle {
    fontName?: string;
    fontSize?: number;
    fontColor?: string;
    bold?: boolean;
    italic?: boolean;
    alignment?: Word.Alignment;
    spaceAfter?: number;
    spaceBefore?: number;
    leftIndent?: number;
}

interface TableStyle {
    fontName?: string;
    fontSize?: number;
    fontColor?: string;
    bold?: boolean;
    italic?: boolean;
}

// 预定义样式常量
const STYLES = {
    // 标题样式
    mainTitle: {
        fontName: 'Times New Roman',
        fontSize: 14,
        fontColor: '#000000',
        bold: true,
        alignment: Word.Alignment.centered,
        spaceAfter: 12,
        spaceBefore: 0
    } as ParagraphStyle,
    
    sectionTitle: {
        fontName: 'Times New Roman',
        fontSize: 14,
        fontColor: '#000000',
        bold: true,
        alignment: Word.Alignment.left,
        spaceAfter: 12,
        spaceBefore: 0,
        italic: false
    } as ParagraphStyle,
    
    tableCaption: {
        fontName: 'Times New Roman',
        fontSize: 12,
        fontColor: '#000000',
        bold: true,
        alignment: Word.Alignment.left,
        spaceAfter: 6,
        spaceBefore: 12,
        italic: false
    } as ParagraphStyle,
    
    // 正文样式
    bodyText: {
        fontName: 'Times New Roman',
        fontSize: 12,
        fontColor: '#000000',
        bold: false,
        italic: false,
        alignment: Word.Alignment.left,
        spaceAfter: 6,
        spaceBefore: 0
    } as ParagraphStyle,
    
    // 页眉页脚样式
    headerFooter: {
        fontName: 'Times New Roman',
        fontSize: 11,
        fontColor: '#000000',
        bold: false,
        italic: false,
        alignment: Word.Alignment.left,
        spaceAfter: 0,
        spaceBefore: 0
    } as ParagraphStyle,
    
    // 脚注样式
    footnote: {
        fontName: 'Times New Roman',
        fontSize: 9,
        fontColor: '#000000',
        bold: false,
        italic: true,
        alignment: Word.Alignment.left,
        spaceAfter: 12,
        spaceBefore: 6
    } as ParagraphStyle,
    
    // TOC样式
    tocEntry: {
        fontName: 'Times New Roman',
        fontSize: 12,
        fontColor: '#0000FF',
        bold: false,
        alignment: Word.Alignment.left,
        leftIndent: 0,
        spaceAfter: 0,
        spaceBefore: 0
    } as ParagraphStyle,
    
    // 表格样式
    tableNormal: {
        fontName: 'Times New Roman',
        fontSize: 9,
        fontColor: '#000000',
        bold: false,
        italic: false
    } as TableStyle,
    
    tableOverview: {
        fontName: 'Times New Roman',
        fontSize: 10,
        fontColor: '#000000',
        bold: false,
        italic: false
    } as TableStyle
};

// ========== 样式应用辅助函数 ==========
function applyParagraphStyle(paragraph: Word.Paragraph, style: ParagraphStyle): void {
    if (style.fontName) paragraph.font.name = style.fontName;
    if (style.fontSize) paragraph.font.size = style.fontSize;
    if (style.fontColor) paragraph.font.color = style.fontColor;
    if (style.bold !== undefined) paragraph.font.bold = style.bold;
    if (style.italic !== undefined) paragraph.font.italic = style.italic;
    if (style.alignment) paragraph.alignment = style.alignment;
    if (style.spaceAfter !== undefined) paragraph.spaceAfter = style.spaceAfter;
    if (style.spaceBefore !== undefined) paragraph.spaceBefore = style.spaceBefore;
    if (style.leftIndent !== undefined) paragraph.leftIndent = style.leftIndent;
}

function applyTableStyle(table: Word.Table, style: TableStyle): void {
    if (style.fontName) table.font.name = style.fontName;
    if (style.fontSize) table.font.size = style.fontSize;
    if (style.fontColor) table.font.color = style.fontColor;
    if (style.bold !== undefined) table.font.bold = style.bold;
    if (style.italic !== undefined) table.font.italic = style.italic;
}

function applyCellStyle(cell: Word.TableCell, style: Partial<ParagraphStyle>, centered: boolean = true): void {
    if (style.fontName) cell.body.font.name = style.fontName;
    if (style.fontSize) cell.body.font.size = style.fontSize;
    if (style.fontColor) cell.body.font.color = style.fontColor;
    if (style.bold !== undefined) cell.body.font.bold = style.bold;
    if (style.italic !== undefined) cell.body.font.italic = style.italic;
    if (centered) {
        cell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
    }
}

// ========== 测试参数定义 ==========
const TEST_PARAMETERS_TABLE2 = [
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

const TEST_PARAMETERS_TABLE2_CONTINUED = [
    { name: "Residue on Ignition (%w/w)", criterion: "≤ 0.2", key: "Residue on Ignition (%w/w)" },
    { name: "Elemental Impurities -- ICP-MS", criterion: "", key: "" },
    { name: "Palladium (ppm)", criterion: "≤ 25", key: "Palladium (ppm)" },
    { name: "Polymorphic Form -- XRPD", criterion: "Conforms to reference standard", key: "Polymorphic Form -- XRPD" },
    { name: "Water Content -- KF (%w/w)", criterion: "Report result", key: "Water Content -- KF (%w/w)" }
];

const TEST_PARAMETERS_TABLE3 = [
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

const ABBREVIATIONS_TEXT = 'Abbreviations: GC = gas chromatography; HPLC = high-performance liquid chromatography; ' +
    'ICP-MS = inductively coupled plasma mass spectrometry; IR = infrared spectroscopy; ' +
    'KF = Karl Fischer; ND = not detected; Pd = Palladium; RRT = relative retention time; ' +
    'XRPD = X‑ray powder diffraction.';

// ========== 文档创建服务类 ==========
export class BatchTemplateService {
    private batchDataList: BatchData[];
    private compoundCode: string;
    
    constructor(batchDataList: BatchData[], compoundCode: string) {
        this.batchDataList = batchDataList;
        this.compoundCode = compoundCode;
    }
    
    // 主入口函数：创建完整文档
    async createCompleteDocument(): Promise<void> {
        await Word.run(async (context) => {
            console.log('📄 开始创建完整的AIMTA文档');
            
            // 清空文档
            context.document.body.clear();
            await context.sync();
            
            // 创建页眉页脚
            try {
                this.createDocumentHeader(context);
                await context.sync();
                console.log('✅ 页眉创建完成');
            } catch (error) {
                console.error('❌ 页眉创建失败:', error);
            }
            
            try {
                this.createDocumentFooter(context);
                await context.sync();
                console.log('✅ 页脚创建完成');
            } catch (error) {
                console.error('❌ 页脚创建失败:', error);
            }
            
            // 创建目录
            try {
                this.createTableOfContentsAndListOfTables(context);
                await context.sync();
                console.log('✅ 目录创建完成');
            } catch (error) {
                console.error('❌ 目录创建失败:', error);
            }
            
            // 插入分页符
            context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
            await context.sync();
            
            // 创建主内容
            try {
                this.createBatchAnalysesTitleAndIntro(context);
                await context.sync();
                console.log('✅ 标题和介绍段落创建完成');
            } catch (error) {
                console.error('❌ 标题创建失败:', error);
            }
            
            // 创建表格
            try {
                await this.createTable1Overview(context);
                await context.sync();
                console.log('✅ 表格1创建完成');
            } catch (error) {
                console.error('❌ 表格1创建失败:', error);
            }
            
            try {
                await this.createTable2BatchAnalysis(context);
                await context.sync();
                console.log('✅ 表格2创建完成');
            } catch (error) {
                console.error('❌ 表格2创建失败:', error);
            }
            
            try {
                await this.createTable3BatchResults(context);
                await context.sync();
                console.log('✅ 表格3创建完成');
            } catch (error) {
                console.error('❌ 表格3创建失败:', error);
            }
            
            console.log('📄 完整AIMTA文档创建完成');
        });
    }
    
    // 创建页眉
    private createDocumentHeader(context: Word.RequestContext): void {
        console.log('📝 创建页眉...');
        
        const section = context.document.sections.getFirst();
        const header = section.getHeader(Word.HeaderFooterType.primary);
        header.clear();
        
        // 第一行
        const firstLine = header.insertParagraph('', Word.InsertLocation.start);
        applyParagraphStyle(firstLine, STYLES.headerFooter);
        firstLine.insertText(this.compoundCode, Word.InsertLocation.start);
        firstLine.insertText('\t\t', Word.InsertLocation.end);
        firstLine.insertText('BeiGene', Word.InsertLocation.end);
        
        // 第二行
        const secondLine = header.insertParagraph(
            `S.4.4 Batch Analyses [${this.compoundCode}, All]`, 
            Word.InsertLocation.end
        );
        applyParagraphStyle(secondLine, {
            ...STYLES.headerFooter,
            spaceAfter: 6
        });
    }
    
    // 创建页脚
    private createDocumentFooter(context: Word.RequestContext): void {
        console.log('📝 创建页脚...');
        
        const section = context.document.sections.getFirst();
        const footer = section.getFooter(Word.HeaderFooterType.primary);
        footer.clear();
        
        const footerParagraph = footer.insertParagraph('', Word.InsertLocation.start);
        applyParagraphStyle(footerParagraph, STYLES.headerFooter);
        
        footerParagraph.insertText('CONFIDENTIAL', Word.InsertLocation.start);
        footerParagraph.insertText('                      ', Word.InsertLocation.end);
        footerParagraph.insertText('Page ', Word.InsertLocation.end);
        
        try {
            const pageRange = footerParagraph.getRange(Word.RangeLocation.end);
            pageRange.insertField(Word.InsertLocation.before, Word.FieldType.page);
            footerParagraph.insertText(' of ', Word.InsertLocation.end);
            const totalPagesRange = footerParagraph.getRange(Word.RangeLocation.end);
            totalPagesRange.insertField(Word.InsertLocation.before, Word.FieldType.numPages);
        } catch (fieldError) {
            console.warn('⚠️ 字段API不可用，使用备选方案');
            footerParagraph.insertText('{ PAGE } of { NUMPAGES }', Word.InsertLocation.end);
        }
    }
    
    // 创建目录
    private createTableOfContentsAndListOfTables(context: Word.RequestContext): void {
        // TOC标题
        const tocTitle = context.document.body.insertParagraph('TABLE OF CONTENTS', Word.InsertLocation.end);
        applyParagraphStyle(tocTitle, STYLES.mainTitle);
        
        // TOC条目
        const tocItems = [
            { text: 'TABLE OF CONTENTS', page: '1' },
            { text: 'LIST OF TABLES', page: '1' },
            { text: 'S.4.4\tBATCH ANALYSES', page: '2' }
        ];
        
        tocItems.forEach(item => {
            const tocEntry = context.document.body.insertParagraph(
                item.text + '\t' + item.page, 
                Word.InsertLocation.end
            );
            applyParagraphStyle(tocEntry, STYLES.tocEntry);
        });
        
        // 空行分隔
        context.document.body.insertParagraph('', Word.InsertLocation.end);
        
        // List of Tables标题
        const listTitle = context.document.body.insertParagraph('LIST OF TABLES', Word.InsertLocation.end);
        applyParagraphStyle(listTitle, {
            ...STYLES.mainTitle,
            spaceBefore: 12
        });
        
        // 表格列表条目
        const tableItems = [
            { text: 'Table 1:\tOverview of BGB-16673 Drug Substance Batches', page: '2' },
            { text: 'Table 2:\tBatch Analysis for Toxicology Batches of BGB-16673 Drug Substance', page: '3' },
            { text: 'Table 3:\tBatch Analysis for GMP Batches of BGB-16673 Drug Substance', page: '5' },
            { text: 'Table 4:\tBatch Results for GMP Batches of BGB-16673 Drug Substance', page: '7' }
        ];
        
        tableItems.forEach(item => {
            const tableEntry = context.document.body.insertParagraph(
                item.text + '\t' + item.page, 
                Word.InsertLocation.end
            );
            applyParagraphStyle(tableEntry, STYLES.tocEntry);
        });
    }
    
    // 创建批次分析标题和介绍
    private createBatchAnalysesTitleAndIntro(context: Word.RequestContext): void {
        // S.4.4 标题
        const sectionTitle = context.document.body.insertParagraph('S.4.4\tBATCH ANALYSES', Word.InsertLocation.end);
        applyParagraphStyle(sectionTitle, STYLES.sectionTitle);
        
        // 介绍段落
        const introPara = context.document.body.insertParagraph(
            `A summary of BGB-16673 drug substance batches manufactured is provided in Table 1. ` +
            `The corresponding batch analysis results for clinical batches are provided in Table 2 and Table 3, respectively.`,
            Word.InsertLocation.end
        );
        applyParagraphStyle(introPara, STYLES.bodyText);
        
        // 说明段落
        const notePara = context.document.body.insertParagraph(
            `The acceptance criteria shown were those effective at the time of testing, unless otherwise specified. ` +
            `Individual relative retention times (RRTs) shown for single unspecified impurities that are variable across ` +
            `batches are shown as ranges, if applicable. Any single unspecified impurities below the reporting threshold ` +
            `(0.05%) across all batches are not reported.`,
            Word.InsertLocation.end
        );
        applyParagraphStyle(notePara, {
            ...STYLES.bodyText,
            spaceAfter: 12
        });
    }
    
    // 创建表格1：Overview
    private async createTable1Overview(context: Word.RequestContext): Promise<void> {
        console.log('📋 开始创建表格1...');
        
        // 表格标题
        const table1Caption = context.document.body.insertParagraph(
            'Table 1:\tOverview of BGB-16673 Drug Substance Batches', 
            Word.InsertLocation.end
        );
        applyParagraphStyle(table1Caption, STYLES.tableCaption);
        
        // 创建表格
        const actualBatchCount = this.batchDataList.length;
        const table1 = context.document.body.insertTable(actualBatchCount + 1, 5, Word.InsertLocation.end);
        table1.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
        applyTableStyle(table1, STYLES.tableOverview);
        
        // 设置表头
        const headers = ['Batch Number', 'Batch Size (kg)', 'Date of Manufacture', 'Manufacturer', 'Use(s)'];
        headers.forEach((header, index) => {
            const cell = table1.getCell(0, index);
            cell.value = header;
            applyCellStyle(cell, { ...STYLES.tableOverview, bold: true });
        });
        
        // 填充数据
        this.batchDataList.forEach((batch, rowIndex) => {
            const manufacturerShort = batch.manufacturer.includes('Changzhou SynTheAll') 
                ? 'Changzhou STA' 
                : batch.manufacturer;
            
            const rowData = [
                batch.batch_number || 'N/A',
                'TBD',
                batch.manufacture_date || 'N/A',
                manufacturerShort,
                'Clinical batch'
            ];
            
            rowData.forEach((data, colIndex) => {
                const cell = table1.getCell(rowIndex + 1, colIndex);
                cell.value = data;
                applyCellStyle(cell, STYLES.tableOverview);
            });
        });
        
        // 表格后空行
        const spacePara = context.document.body.insertParagraph('', Word.InsertLocation.end);
        spacePara.spaceAfter = 12;
        spacePara.spaceBefore = 6;
    }
    
    // 创建表格2：Batch Analysis
    private async createTable2BatchAnalysis(context: Word.RequestContext): Promise<void> {
        // 表格标题
        const table2Caption = context.document.body.insertParagraph(
            'Table 2: Batch Analysis for GMP Batches of BGB-16673 Drug Substance', 
            Word.InsertLocation.end
        );
        applyParagraphStyle(table2Caption, STYLES.tableCaption);
        
        // 创建表格
        const actualBatchCount = this.batchDataList.length;
        const table2 = context.document.body.insertTable(
            TEST_PARAMETERS_TABLE2.length + 1, 
            actualBatchCount + 2, 
            Word.InsertLocation.end
        );
        table2.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
        applyTableStyle(table2, STYLES.tableNormal);
        
        // 设置表头
        this.setupTable2Headers(table2, actualBatchCount);
        
        // 填充测试参数数据
        this.fillTestParameters(table2, TEST_PARAMETERS_TABLE2, 1);
        
        // 表格后空行
        context.document.body.insertParagraph('', Word.InsertLocation.end);
        
        // 创建继续表格部分
        await this.createTable2Continued(context);
    }
    
    // 创建表格2继续部分
    private async createTable2Continued(context: Word.RequestContext): Promise<void> {
        // 继续表格标题
        const continuedCaption = context.document.body.insertParagraph(
            'Table 2: Batch Analysis for GMP Batches of BGB-16673 Drug Substance (Continued)', 
            Word.InsertLocation.end
        );
        applyParagraphStyle(continuedCaption, STYLES.tableCaption);
        
        // 创建继续表格
        const actualBatchCount = this.batchDataList.length;
        const continuedTable = context.document.body.insertTable(
            TEST_PARAMETERS_TABLE2_CONTINUED.length + 1, 
            actualBatchCount + 2, 
            Word.InsertLocation.end
        );
        continuedTable.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
        applyTableStyle(continuedTable, STYLES.tableNormal);
        
        // 设置表头
        this.setupTable2Headers(continuedTable, actualBatchCount);
        
        // 填充测试参数数据
        this.fillTestParameters(continuedTable, TEST_PARAMETERS_TABLE2_CONTINUED, 1);
        
        // 添加缩写说明
        const abbreviations = context.document.body.insertParagraph(
            ABBREVIATIONS_TEXT,
            Word.InsertLocation.end
        );
        applyParagraphStyle(abbreviations, STYLES.footnote);
        
        // 表格后空行
        const spacePara = context.document.body.insertParagraph('', Word.InsertLocation.end);
        spacePara.spaceAfter = 12;
    }
    
    // 创建表格3：Batch Results
    private async createTable3BatchResults(context: Word.RequestContext): Promise<void> {
        // 表格标题
        const table3Caption = context.document.body.insertParagraph(
            'Table 3: Batch Results for GMP Batches of BGB-16673 Drug Substance', 
            Word.InsertLocation.end
        );
        applyParagraphStyle(table3Caption, STYLES.tableCaption);
        
        // 使用最新批次作为示例
        const selectedBatch = this.batchDataList.length > 0 
            ? this.batchDataList[this.batchDataList.length - 1] 
            : null;
        
        if (!selectedBatch) {
            console.warn('没有批次数据可用于表格3');
            return;
        }
        
        // 创建表格3
        const table3 = context.document.body.insertTable(
            TEST_PARAMETERS_TABLE3.length + 1, 
            3, 
            Word.InsertLocation.end
        );
        table3.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
        applyTableStyle(table3, STYLES.tableNormal);
        
        // 设置表头
        const headers3 = ['Test Parameter', 'Acceptance Criterion', selectedBatch.batch_number || 'Latest Batch'];
        headers3.forEach((header, index) => {
            const cell = table3.getCell(0, index);
            cell.value = header;
            applyCellStyle(cell, { ...STYLES.tableNormal, bold: true });
        });
        
        // 填充表格3数据
        TEST_PARAMETERS_TABLE3.forEach((param, rowIndex) => {
            const paramCell = table3.getCell(rowIndex + 1, 0);
            paramCell.value = param.name;
            if (param.name.includes('--')) {
                applyCellStyle(paramCell, { ...STYLES.tableNormal, bold: true }, false);
            } else {
                applyCellStyle(paramCell, STYLES.tableNormal, false);
            }
            
            const criterionCell = table3.getCell(rowIndex + 1, 1);
            criterionCell.value = param.criterion;
            applyCellStyle(criterionCell, STYLES.tableNormal, false);
            
            const resultCell = table3.getCell(rowIndex + 1, 2);
            let result = '';
            
            if (param.key && selectedBatch.test_results && selectedBatch.test_results[param.key]) {
                result = selectedBatch.test_results[param.key];
            } else if (!param.key || param.criterion === '') {
                result = '';
            } else {
                result = 'TBD';
            }
            
            resultCell.value = result;
            applyCellStyle(resultCell, STYLES.tableNormal);
        });
        
        // 添加缩写说明
        const abbreviations3 = context.document.body.insertParagraph(
            ABBREVIATIONS_TEXT,
            Word.InsertLocation.end
        );
        applyParagraphStyle(abbreviations3, STYLES.footnote);
    }
    
    // ========== 辅助方法 ==========
    
    // 设置表格2的表头
    private setupTable2Headers(table: Word.Table, batchCount: number): void {
        const headerCell1 = table.getCell(0, 0);
        headerCell1.value = 'Test Parameter';
        applyCellStyle(headerCell1, { ...STYLES.tableNormal, bold: true });
        
        const headerCell2 = table.getCell(0, 1);
        headerCell2.value = 'Acceptance Criterion';
        applyCellStyle(headerCell2, { ...STYLES.tableNormal, bold: true });
        
        // 批次号表头
        this.batchDataList.forEach((batch, index) => {
            const headerCell = table.getCell(0, index + 2);
            headerCell.value = batch.batch_number || `Batch ${index + 1}`;
            applyCellStyle(headerCell, { ...STYLES.tableNormal, bold: true });
        });
    }
    
    // 填充测试参数数据
    private fillTestParameters(
        table: Word.Table, 
        parameters: Array<{name: string, criterion: string, key: string}>,
        startRow: number
    ): void {
        parameters.forEach((param, rowIndex) => {
            // 测试参数名称
            const paramCell = table.getCell(rowIndex + startRow, 0);
            paramCell.value = param.name;
            if (param.name.includes('--')) {
                applyCellStyle(paramCell, { ...STYLES.tableNormal, bold: true });
            } else {
                applyCellStyle(paramCell, STYLES.tableNormal);
            }
            
            // 接受标准
            const criterionCell = table.getCell(rowIndex + startRow, 1);
            criterionCell.value = param.criterion;
            applyCellStyle(criterionCell, STYLES.tableNormal);
            
            // 批次结果
            this.batchDataList.forEach((batch, batchIndex) => {
                const resultCell = table.getCell(rowIndex + startRow, batchIndex + 2);
                let result = '';
                
                if (param.key && batch.test_results && batch.test_results[param.key]) {
                    result = batch.test_results[param.key];
                } else if (!param.key || param.criterion === '') {
                    result = ''; // 空白行
                } else {
                    result = 'TBD'; // 未找到数据时显示TBD
                }
                
                resultCell.value = result;
                applyCellStyle(resultCell, STYLES.tableNormal);
            });
        });
    }
}

// ========== 导出便捷函数 ==========

/**
 * 创建完整的AIMTA文档
 * @param batchDataList 批次数据列表
 * @param compoundCode 化合物代码
 * @returns Promise<void>
 */
export async function createAIMTADocument(
    batchDataList: BatchData[], 
    compoundCode: string
): Promise<void> {
    const service = new BatchTemplateService(batchDataList, compoundCode);
    return service.createCompleteDocument();
}

/**
 * 获取样式配置（用于调试或自定义）
 * @returns 样式配置对象
 */
export function getStyleConfigs() {
    return STYLES;
}

/**
 * 导出样式应用函数（如果需要在其他地方使用）
 */
export { applyParagraphStyle, applyTableStyle, applyCellStyle };