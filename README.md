# 學生成績顯示 Web Part

這是一個使用SharePoint Framework (SPFx) 開發的學生成績顯示Web Part，使用純JavaScript/TypeScript實現，無需額外的框架依賴。

## 功能特點

- 📊 **動態成績表格** - 顯示學生的完整成績信息
- 🔍 **搜索功能** - 按學生姓名或學號搜索
- 🎯 **科目篩選** - 按科目篩選（80分以上）
- 📈 **統計分析** - 實時計算總學生數、平均分、最高分、最低分
- 📥 **數據導出** - 支持CSV格式導出
- 📱 **響應式設計** - 適配各種屏幕尺寸
- 🌙 **主題支持** - 支持淺色和深色主題
- 🎨 **現代化UI** - 使用Office UI Fabric設計語言

## 技術棧

- SharePoint Framework (SPFx)
- TypeScript
- SCSS/CSS Modules
- 純JavaScript DOM操作

## 開發環境要求

- Node.js 18.17.1 或更高版本
- npm 9.6.7 或更高版本
- SharePoint Framework Yeoman生成器

## 安裝和運行

### 1. 安裝依賴
```bash
npm install
```

### 2. 本地開發
```bash
npm run serve
```

### 3. 構建項目
```bash
npm run build
```

### 4. 打包解決方案
```bash
npm run package-solution
```

## 數據結構

Web Part使用以下學生數據結構：

```typescript
interface Student {
  studentId: string;    // 學號
  name: string;         // 姓名
  chinese: number;      // 中文成績
  english: number;      // 英文成績
  math: number;         // 數學成績
  science: number;      // 科學成績
  total: number;        // 總分
  average: number;      // 平均分
}
```

## 模擬數據

目前使用模擬數據，包含10個學生的完整成績信息：

- 張小明 (S001) - 中文:85, 英文:92, 數學:88, 科學:90
- 李小華 (S002) - 中文:78, 英文:85, 數學:92, 科學:87
- 王小美 (S003) - 中文:92, 英文:88, 數學:85, 科學:94
- 陳小強 (S004) - 中文:80, 英文:90, 數學:95, 科學:82
- 林小芳 (S005) - 中文:88, 英文:85, 數學:90, 科學:88
- 黃小偉 (S006) - 中文:75, 英文:92, 數學:87, 科學:90
- 劉小玲 (S007) - 中文:90, 英文:88, 數學:93, 科學:85
- 吳小傑 (S008) - 中文:82, 英文:85, 數學:88, 科學:92
- 趙小雅 (S009) - 中文:95, 英文:90, 數學:89, 科學:91
- 孫小龍 (S010) - 中文:87, 英文:93, 數學:91, 科學:89

## 功能說明

### 搜索功能
- 在搜索框中輸入學生姓名或學號
- 支持實時搜索，無需點擊按鈕
- 搜索結果會即時更新表格和統計信息

### 科目篩選
- 選擇特定科目（中文、英文、數學、科學）
- 只顯示該科目成績80分以上的學生
- 可以與搜索功能組合使用

### 統計信息
- **總學生數** - 當前顯示的學生數量
- **平均分數** - 當前學生的平均分數
- **最高分數** - 當前學生中的最高平均分
- **最低分數** - 當前學生中的最低平均分

### 數據導出
- 點擊"匯出 CSV"按鈕
- 自動下載包含當前篩選結果的CSV文件
- 文件名格式：`student_grades.csv`

## 部署到SharePoint

### 方法1: 開發環境部署
1. 運行 `npm run serve`
2. 在SharePoint頁面中添加Web Part
3. 配置適當的權限

### 方法2: 生產環境部署
1. 運行 `npm run package-solution`
2. 將生成的 `.sppkg` 文件上傳到SharePoint App Catalog
3. 在SharePoint頁面中添加Web Part

## 自定義配置

### 修改模擬數據
在 `ScoreDisplayWebPart.ts` 文件的 `_loadStudentData()` 方法中修改學生數據：

```typescript
this.students = [
  // 添加或修改學生數據
  { studentId: 'S011', name: '新學生', chinese: 90, english: 85, math: 88, science: 92, total: 355, average: 88.75 },
  // ...
];
```

### 連接真實數據源
將 `_loadStudentData()` 方法中的模擬數據替換為SharePoint API調用：

```typescript
private async _loadStudentData(): Promise<void> {
  try {
    const response = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('StudentGrades')/items`,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();
    this.students = data.value.map(item => ({
      studentId: item.StudentID,
      name: item.Name,
      chinese: item.Chinese,
      english: item.English,
      math: item.Math,
      science: item.Science,
      total: item.Total,
      average: item.Average
    }));
  } catch (error) {
    console.error('Error loading student data:', error);
  }
}
```

## 樣式自定義

Web Part使用CSS Modules，樣式文件位於：
- `ScoreDisplayWebPart.module.scss` - 主要樣式文件

支持的主題變體：
- 淺色主題（默認）
- 深色主題（自動適配）

## 故障排除

### 常見問題
1. **Web Part不顯示** - 檢查瀏覽器控制台錯誤
2. **樣式問題** - 確保CSS Modules正確編譯
3. **數據不載入** - 檢查模擬數據或API調用
4. **搜索不工作** - 檢查事件監聽器設置

### 調試技巧
- 使用瀏覽器開發者工具檢查DOM結構
- 查看控制台錯誤信息
- 檢查網絡請求（如果使用API）

## 未來改進

- [ ] 添加排序功能
- [ ] 支持分頁顯示
- [ ] 添加圖表視圖
- [ ] 支持批量操作
- [ ] 添加權限控制
- [ ] 支持多語言

## 貢獻

歡迎提交Issue和Pull Request來改進這個Web Part。

## 授權

此項目使用MIT授權。
