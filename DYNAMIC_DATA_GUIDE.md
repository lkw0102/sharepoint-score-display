# 動態數據功能說明 - 學生成績顯示 Web Part

## 🎯 功能概述

新的Web Part版本現在能夠根據不同的SharePoint頁面和當前用戶信息動態加載不同的學生成績數據。

## 🔍 核心功能

### 1. 頁面信息獲取
Web Part會自動獲取當前SharePoint頁面的信息：
- **頁面名稱**：從URL中提取（例如：TEST1, TEST2, ClassA等）
- **站點URL**：當前SharePoint站點的完整URL
- **列表標題**：自動生成對應的SharePoint列表名稱
- **Excel文件名**：自動生成對應的Excel文件名

### 2. 用戶信息獲取
Web Part會獲取當前登錄用戶的信息：
- **顯示名稱**：用戶的顯示名稱
- **電子郵件**：用戶的電子郵件地址
- **登錄名稱**：用戶的登錄名稱
- **管理員權限**：是否為站點管理員

### 3. 動態數據加載
根據頁面名稱自動加載對應的學生成績數據：

#### 支持的頁面類型：
- **TEST1**：測試班級1的學生數據
- **TEST2**：測試班級2的學生數據
- **ClassA**：A班的學生數據
- **ClassB**：B班的學生數據
- **其他頁面**：默認學生數據

## 📊 數據結構

### 學生數據格式
```typescript
interface Student {
  studentId: string;    // 學號（包含班級前綴）
  name: string;         // 學生姓名
  chinese: number;      // 中文成績
  english: number;      // 英文成績
  math: number;         // 數學成績
  science: number;      // 科學成績
  total: number;        // 總分
  average: number;      // 平均分
}
```

### 頁面信息格式
```typescript
interface PageInfo {
  pageName: string;     // 頁面名稱
  siteUrl: string;      // 站點URL
  listTitle?: string;   // SharePoint列表標題
  excelFileName?: string; // Excel文件名
}
```

### 用戶信息格式
```typescript
interface UserInfo {
  displayName: string;  // 顯示名稱
  email: string;        // 電子郵件
  loginName: string;    // 登錄名稱
  isSiteAdmin: boolean; // 是否為管理員
}
```

## 🚀 使用示例

### 1. 訪問TEST1頁面
**URL：** `https://groupespauloedu.sharepoint.com/sites/Classrooms/SitePages/TEST1.aspx`

**顯示信息：**
- 當前頁面：TEST1 (https://groupespauloedu.sharepoint.com/sites/Classrooms)
- 當前用戶：[用戶名稱] ([用戶郵箱])

**加載數據：**
- 學號：T1-001, T1-002, T1-003
- 學生：張小明, 李小華, 王小美

### 2. 訪問ClassA頁面
**URL：** `https://groupespauloedu.sharepoint.com/sites/Classrooms/SitePages/ClassA.aspx`

**加載數據：**
- 學號：CA-001, CA-002, CA-003
- 學生：劉小玲, 吳小傑, 趙小雅

## 🔧 技術實現

### 1. 頁面信息獲取
```typescript
private async _getPageInfo(): Promise<PageInfo> {
  const currentUrl = window.location.href;
  const urlParts = currentUrl.split('/');
  const pageName = urlParts[urlParts.length - 1].replace('.aspx', '');
  
  return {
    pageName: pageName,
    siteUrl: this.context.pageContext.web.absoluteUrl,
    listTitle: `成績資料_${pageName}`,
    excelFileName: `學生成績_${pageName}.xlsx`
  };
}
```

### 2. 用戶信息獲取
```typescript
private async _getUserInfo(): Promise<UserInfo> {
  const response = await this.context.spHttpClient.get(
    `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser`,
    SPHttpClient.configurations.v1
  );
  
  const userData = await response.json();
  return {
    displayName: userData.Title,
    email: userData.Email,
    loginName: userData.LoginName,
    isSiteAdmin: userData.IsSiteAdmin
  };
}
```

### 3. 動態數據加載
```typescript
private _getMockDataByPage(pageName: string): Student[] {
  const mockDataMap: { [key: string]: Student[] } = {
    'TEST1': [/* TEST1的學生數據 */],
    'TEST2': [/* TEST2的學生數據 */],
    'ClassA': [/* ClassA的學生數據 */],
    'ClassB': [/* ClassB的學生數據 */]
  };
  
  return mockDataMap[pageName] || [/* 默認數據 */];
}
```

## 📁 文件結構

```
student-grades-spfx-new/
├── src/webparts/scoreDisplay/
│   ├── ScoreDisplayWebPart.ts          # 主要邏輯文件
│   ├── ScoreDisplayWebPart.module.scss # 樣式文件
│   └── ScoreDisplayWebPart.manifest.json # 配置清單
├── sharepoint/solution/
│   └── student-grades-display.sppkg    # 部署包
└── DYNAMIC_DATA_GUIDE.md               # 本說明文檔
```

## 🔄 擴展功能

### 1. 添加新的虛擬課室
在 `_getMockDataByPage` 方法中添加新的數據映射：

```typescript
'ClassC': [
  { studentId: 'CC-001', name: '新學生1', chinese: 85, english: 90, math: 88, science: 92, total: 355, average: 88.75 },
  { studentId: 'CC-002', name: '新學生2', chinese: 88, english: 85, math: 90, science: 87, total: 350, average: 87.5 }
]
```

### 2. 連接真實SharePoint列表
將模擬數據替換為真實的SharePoint REST API調用：

```typescript
private async _loadStudentDataFromSharePoint(listTitle: string): Promise<Student[]> {
  const response = await this.context.spHttpClient.get(
    `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`,
    SPHttpClient.configurations.v1
  );
  
  const data = await response.json();
  return data.value.map(item => ({
    studentId: item.StudentId,
    name: item.Name,
    chinese: item.Chinese,
    english: item.English,
    math: item.Math,
    science: item.Science,
    total: item.Total,
    average: item.Average
  }));
}
```

### 3. 連接Excel文件
使用Microsoft Graph API或SharePoint REST API讀取Excel文件：

```typescript
private async _loadStudentDataFromExcel(fileName: string): Promise<Student[]> {
  // 使用Microsoft Graph API讀取Excel文件
  // 或使用SharePoint REST API
}
```

## 🎨 用戶界面

### 頁面信息顯示
Web Part會在標題區域顯示：
- 當前頁面名稱和URL
- 當前用戶名稱和郵箱

### 動態文件名
CSV導出功能會根據頁面名稱生成對應的文件名：
- TEST1頁面：`student_grades_TEST1.csv`
- ClassA頁面：`student_grades_ClassA.csv`

## 🔒 權限和安全

### 用戶權限檢查
Web Part會檢查用戶的權限：
- 讀取權限：所有用戶都可以查看成績
- 管理員權限：可以進行額外操作

### 數據安全
- 所有用戶輸入都經過HTML轉義
- API調用使用SharePoint認證
- 錯誤信息不會暴露敏感數據

## 🚀 部署說明

### 1. 上傳新版本
1. 進入SharePoint管理中心
2. 應用程式目錄 → 上傳
3. 選擇 `student-grades-display.sppkg`
4. 點擊部署

### 2. 測試不同頁面
1. 訪問 `https://groupespauloedu.sharepoint.com/sites/Classrooms/SitePages/TEST1.aspx`
2. 添加Web Part並驗證數據
3. 訪問其他頁面測試動態加載

### 3. 監控和調試
- 使用瀏覽器開發者工具查看控制台日誌
- 檢查網絡請求是否成功
- 驗證頁面信息是否正確獲取

## 📞 支持和維護

### 常見問題
1. **頁面信息不顯示**：檢查URL格式是否正確
2. **用戶信息獲取失敗**：檢查用戶權限
3. **數據加載失敗**：檢查網絡連接和API權限

### 聯繫支持
如果遇到問題，請提供：
- 頁面URL
- 錯誤信息截圖
- 瀏覽器控制台日誌
- 用戶權限信息

---

**注意：** 此版本支持動態數據加載，可以根據不同的虛擬課室頁面顯示對應的學生成績數據。未來可以擴展為連接真實的SharePoint列表或Excel文件。
