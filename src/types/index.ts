// Compound types
export interface Compound {
    id: string;
    code: string;
    name: string;
    description?: string;
}

// Template types
export interface Template {
    id: string;
    compoundId: string;
    region: 'CN' | 'EU' | 'US';
    templateContent: string;
    fieldMapping?: Record<string, string>;
}

// COA Document types
export interface COADocument {
    id: string;
    compoundId: string;
    filename: string;
    filePath: string;
    processingStatus: 'pending' | 'processing' | 'completed' | 'failed';
    uploadedAt: Date;
    processedAt?: Date;
}

// Batch Analysis Data types
export interface BatchData {
    filename: string;
    batch_number: string;
    manufacture_date: string;
    manufacturer: string;
    test_results: Record<string, string>;
}

// Extracted data types (legacy)
export interface ExtractedField {
    id: string;
    documentId?: string;
    fieldName: string;
    fieldValue: string;
    confidenceScore: number;
    originalText?: string;
    filename?: string;
}

// API Response types
export interface ApiResponse<T> {
    success: boolean;
    data?: T;
    error?: string;
    message?: string;
}

// Directory processing request
export interface DirectoryProcessRequest {
    compound_id: string;
    template_id: string;
    directory_path?: string;
}

// Batch processing result
export interface BatchProcessingResult {
    processedFiles: string[];
    failedFiles: Array<{filename: string; error: string}>;
    totalFiles: number;
    batchData: BatchData[];
    status: 'success' | 'partial' | 'failed';
    message?: string;
}

// App state types (updated for batch processing)
export interface AppState {
    selectedCompound?: Compound;
    selectedTemplate?: Template;
    extractedData: ExtractedField[]; // Legacy support
    isLoading: boolean;
    error?: string;
}

// Test Parameters from base template
export interface TestParameter {
    name: string;
    acceptanceCriterion: string;
    key: string;
}

// COA Table Structure
export interface COATable {
    title: string;
    headers: string[];
    parameters: TestParameter[];
    batchData: BatchData[];
}

// API Endpoints
export const API_BASE_URL = (() => {
  const hostname = window.location.hostname;

  console.log(`前端域名: ${hostname}`);
  console.log(`完整URL: ${window.location.href}`);  

  // 生产环境 - beone-d.beigenecorp.net
  if (hostname === "beone-d.beigenecorp.net" || hostname.includes("beigenecorp.net")) {
    const apiUrl = "https://beone-d.beigenecorp.net/api/aimta";
    console.log(`使用生产环境API: ${apiUrl}`);
    return apiUrl;
  }
  
  // 本地开发环境
  if (hostname === "localhost" || hostname === "127.0.0.1") {
    const apiUrl = "https://localhost:8000";
    console.log(`使用本地开发API: ${apiUrl}`);
    return apiUrl;
  }
  
  // 开发环境IP
  if (hostname === "10.8.63.207") {
    const apiUrl = "https://10.8.63.207:8000";
    console.log(`使用开发IP API: ${apiUrl}`);
    return apiUrl;
  }
  
  // 默认返回生产环境
  const defaultUrl = "https://beone-d.beigenecorp.net/api/aimta";
  console.log(`使用默认API: ${defaultUrl}`);
  return defaultUrl;
})();


// 测试API连接的辅助函数
export const testApiConnection = async (): Promise<boolean> => {
  try {
    console.log(`🔍 测试API连接: ${API_BASE_URL}`);
    
    const response = await fetch(`${API_BASE_URL}/api/health`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
      },
      // 添加CORS相关配置
      mode: 'cors',
      credentials: 'omit'
    });
    
    if (response.ok) {
      const data = await response.json();
      console.log(`✅ API连接成功:`, data);
      return true;
    } else {
      console.error(`❌ API连接失败: ${response.status} ${response.statusText}`);
      return false;
    }
  } catch (error) {
    console.error(`❌ API连接异常:`, error);
    return false;
  }
};

// Template test parameters mapping
export const TEMPLATE_TEST_PARAMETERS = {
    TABLE_2: [
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
    ],
    TABLE_3: [
        { name: "Residue on Ignition (%w/w)", criterion: "≤ 0.2", key: "Residue on Ignition (%w/w)" },
        { name: "Elemental Impurities -- ICP-MS", criterion: "", key: "" },
        { name: "Palladium (ppm)", criterion: "≤ 25", key: "Palladium (ppm)" },
        { name: "Polymorphic Form -- XRPD", criterion: "Conforms to reference standard", key: "Polymorphic Form -- XRPD" },
        { name: "Water Content -- KF (%w/w)", criterion: "Report result", key: "Water Content -- KF (%w/w)" },
        { name: "RRT 0.83", criterion: "", key: "RRT 0.83" }
    ]
};