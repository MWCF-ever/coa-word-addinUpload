// src/types/index.ts - 修复Mixed Content问题

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

// 🔥 修复：更智能的API_BASE_URL检测，解决Mixed Content问题
export const API_BASE_URL = (() => {
  const hostname = window.location.hostname;
  const protocol = window.location.protocol;
  const href = window.location.href;

  console.log(`🔍 API_BASE_URL检测:`);
  console.log(`  - hostname: ${hostname}`);
  console.log(`  - protocol: ${protocol}`);
  console.log(`  - href: ${href}`);

  // 🔥 关键修复：Office 365 环境检测
  if (hostname.includes("sharepoint.com") || 
      hostname.includes("office.com") ||
      hostname.includes("officeapps.live.com") ||
      hostname.includes("outlook.office") ||
      href.includes("sharepoint.com")) {
    
    const apiUrl = "https://beone-d.beigenecorp.net/api/aimta";
    console.log(`✅ 检测到Office 365环境，使用生产API: ${apiUrl}`);
    return apiUrl;
  }

  // 生产环境 - beone-d.beigenecorp.net
  if (hostname === "beone-d.beigenecorp.net" || hostname.includes("beigenecorp.net")) {
    const apiUrl = "https://beone-d.beigenecorp.net/api/aimta";
    console.log(`✅ 使用生产环境API: ${apiUrl}`);
    return apiUrl;
  }
  
  // 本地开发环境
  if (hostname === "localhost" || hostname === "127.0.0.1") {
    const apiUrl = "https://localhost:8000";
    console.log(`✅ 使用本地开发API: ${apiUrl}`);
    return apiUrl;
  }
  
  // 开发环境IP
  if (hostname === "10.8.63.207") {
    const apiUrl = "https://10.8.63.207:8000";
    console.log(`✅ 使用开发IP API: ${apiUrl}`);
    return apiUrl;
  }
  
  // 🔥 默认强制使用HTTPS生产环境，避免Mixed Content
  const defaultUrl = "https://beone-d.beigenecorp.net/api/aimta";
  console.log(`✅ 使用默认HTTPS API: ${defaultUrl}`);
  return defaultUrl;
})();

// 增强的API连接测试函数
export const testApiConnection = async (): Promise<{success: boolean, responseTime: number, error?: string}> => {
  const startTime = Date.now();
  
  try {
    console.log(`🔍 测试API连接: ${API_BASE_URL}`);
    
    // 🔥 修复：确保使用HTTPS，避免Mixed Content
    const testUrl = API_BASE_URL.replace('http://', 'https://');
    
    // 尝试多个端点
    const endpoints = ['/health', '/', '/docs'];
    let lastError = null;
    
    for (const endpoint of endpoints) {
      try {
        console.log(`📡 测试端点: ${testUrl}${endpoint}`);
        
        const response = await fetch(`${testUrl}${endpoint}`, {
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          mode: 'cors',
          credentials: 'omit'
        });
        
        const responseTime = Date.now() - startTime;
        
        if (response.ok) {
          try {
            const data = await response.json();
            console.log(`✅ API连接成功 (${responseTime}ms) 端点: ${endpoint}`, data);
            return { success: true, responseTime };
          } catch (jsonError) {
            // 如果不是JSON响应，但状态码是200，也算成功
            console.log(`✅ API连接成功 (${responseTime}ms) 端点: ${endpoint} (非JSON响应)`);
            return { success: true, responseTime };
          }
        } else {
          lastError = `HTTP ${response.status}: ${response.statusText}`;
          console.log(`❌ 端点 ${endpoint} 失败: ${lastError}`);
        }
      } catch (endpointError: any) {
        lastError = endpointError.message;
        console.log(`❌ 端点 ${endpoint} 异常: ${lastError}`);
      }
    }
    
    // 所有端点都失败了
    const responseTime = Date.now() - startTime;
    console.error(`❌ 所有API端点都失败了，最后一个错误: ${lastError}`);
    return { success: false, responseTime, error: lastError || 'All endpoints failed' };
    
  } catch (error: any) {
    const responseTime = Date.now() - startTime;
    console.error(`❌ API连接异常:`, error);
    
    let errorMessage = 'Unknown error';
    if (error.name === 'TypeError' && error.message.includes('Failed to fetch')) {
      errorMessage = 'Network connection failed - CORS or server unreachable';
    } else if (error.message.includes('ERR_CONNECTION_REFUSED')) {
      errorMessage = 'Connection refused - server not running';
    } else {
      errorMessage = error.message;
    }
    
    return { success: false, responseTime, error: errorMessage };
  }
};

// 调试用API测试函数
export const debugApiEndpoints = async (): Promise<void> => {
  console.log('🔧 开始API端点调试...');
  
  const testUrls = [
    `${API_BASE_URL}`,
    `${API_BASE_URL}/`,
    `${API_BASE_URL}/health`,
    `${API_BASE_URL}/compounds`,
    `${API_BASE_URL}/docs`,
    // 🔥 修复：强制HTTPS测试
    'https://beone-d.beigenecorp.net/api/aimta/health',
    'https://beone-d.beigenecorp.net/api/aimta/',
  ];
  
  for (const url of testUrls) {
    try {
      console.log(`\n📡 测试: ${url}`);
      const response = await fetch(url, {
        method: 'GET',
        mode: 'cors',
        credentials: 'omit'
      });
      
      console.log(`   状态: ${response.status} ${response.statusText}`);
      console.log(`   头部: Content-Type=${response.headers.get('content-type')}`);
      
      if (response.ok) {
        try {
          const data = await response.json();
          console.log(`   响应:`, data);
        } catch {
          const text = await response.text();
          console.log(`   文本响应:`, text.substring(0, 200));
        }
      }
    } catch (error: any) {
      console.log(`   错误: ${error.message}`);
    }
  }
  
  console.log('\n🔧 API端点调试完成');
};

// 可以在浏览器控制台调用的全局调试函数
declare global {
  interface Window {
    debugApiEndpoints: () => Promise<void>;
    testApiConnection: () => Promise<{success: boolean, responseTime: number, error?: string}>;
    getCurrentApiBaseUrl: () => string;
  }
}

// 暴露到全局，方便调试
if (typeof window !== 'undefined') {
  window.debugApiEndpoints = debugApiEndpoints;
  window.testApiConnection = testApiConnection;
  window.getCurrentApiBaseUrl = () => API_BASE_URL;
}

// 快速修复建议检查函数
export const quickFixSuggestions = () => {
  console.log(`
🔧 Mixed Content错误快速修复建议:

当前检测结果:
- 当前域名: ${window.location.hostname}
- 当前协议: ${window.location.protocol}
- API地址: ${API_BASE_URL}

问题诊断:
1. Mixed Content错误通常由HTTP/HTTPS混合引起
2. Office 365环境强制HTTPS
3. 需要确保所有API调用都使用HTTPS

修复步骤:
1. 确认后端API支持HTTPS
2. 更新nginx配置支持SSL
3. 确保所有API调用使用HTTPS URL

调试命令:
- window.getCurrentApiBaseUrl() - 查看当前API地址
- window.testApiConnection() - 测试API连接
- window.debugApiEndpoints() - 测试所有端点
  `);
};

// 自动在控制台显示修复建议
if (typeof window !== 'undefined') {
  console.log('🔧 API调试工具已加载 (Mixed Content修复版)');
  console.log('运行 quickFixSuggestions() 获取修复建议');
  console.log(`当前API地址: ${API_BASE_URL}`);
  (window as any).quickFixSuggestions = quickFixSuggestions;
}

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