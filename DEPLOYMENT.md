# 學生成績顯示 Web Part 部署指南

## 🎉 恭喜！你的SPFx Web Part已經成功打包

### 📦 生成的文件

打包過程已成功生成以下文件：
- **`sharepoint/solution/student-grades-display.sppkg`** - 這是可以部署到SharePoint的解決方案包

## 🚀 部署到SharePoint

### 方法1: 使用SharePoint App Catalog（推薦）

#### 步驟1: 上傳解決方案包
1. 登入你的SharePoint管理員帳戶
2. 進入 **SharePoint 管理中心**
3. 點擊 **應用程式目錄** (App Catalog)
4. 選擇你的租用戶應用程式目錄
5. 點擊 **上傳** 按鈕
6. 選擇 `student-grades-display.sppkg` 文件
7. 點擊 **部署**

#### 步驟2: 在SharePoint頁面中使用
1. 進入任何SharePoint網站
2. 編輯頁面
3. 點擊 **"+"** 添加Web Part
4. 在搜索框中輸入 **"學生成績顯示"**
5. 選擇並添加Web Part
6. 配置Web Part屬性（可選）
7. 發布頁面

### 方法2: 使用SharePoint Online PowerShell

```powershell
# 連接到SharePoint Online
Connect-SPOService -Url https://yourtenant-admin.sharepoint.com

# 上傳應用程式包
Add-SPOApp -Path "C:\path\to\student-grades-display.sppkg" -Overwrite

# 部署應用程式
Install-SPOApp -Identity "student-grades-display-client-side-solution" -Web "https://yourtenant.sharepoint.com/sites/yoursite"
```

### 方法3: 使用SharePoint Framework CLI

```bash
# 安裝SPFx CLI工具
npm install -g @microsoft/spfx-cli

# 部署到SharePoint
spfx package deploy --path sharepoint/solution/student-grades-display.sppkg --tenant yourtenant.sharepoint.com
```

## 🔧 開發環境測試

如果你想在開發環境中測試Web Part：

```bash
# 啟動開發服務器
npm run serve

# 或者使用gulp
npx gulp serve
```

然後訪問：`https://yourtenant.sharepoint.com/_layouts/workbench.aspx`

## 📋 功能驗證清單

部署完成後，請驗證以下功能：

### ✅ 基本功能
- [ ] Web Part正常顯示
- [ ] 學生成績表格載入
- [ ] 搜索功能正常
- [ ] 科目篩選正常
- [ ] 統計信息正確
- [ ] CSV導出功能正常

### ✅ 響應式設計
- [ ] 桌面端顯示正常
- [ ] 平板端顯示正常
- [ ] 手機端顯示正常

### ✅ 主題支持
- [ ] 淺色主題正常
- [ ] 深色主題正常

## 🛠️ 自定義配置

### 修改學生數據
在 `src/webparts/scoreDisplay/ScoreDisplayWebPart.ts` 中修改 `_loadStudentData()` 方法：

```typescript
private _loadStudentData(): void {
  this.students = [
    // 添加你的學生數據
    { studentId: 'S001', name: '學生姓名', chinese: 85, english: 92, math: 88, science: 90, total: 355, average: 88.75 },
    // ...
  ];
}
```

### 連接SharePoint List
將模擬數據替換為真實的SharePoint API調用：

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

## 🔍 故障排除

### 常見問題

1. **Web Part不顯示**
   - 檢查瀏覽器控制台錯誤
   - 確認解決方案包已正確部署
   - 檢查用戶權限

2. **數據不載入**
   - 檢查模擬數據是否正確
   - 如果使用API，檢查網絡請求
   - 確認SharePoint List權限

3. **樣式問題**
   - 清除瀏覽器緩存
   - 檢查CSS是否正確編譯
   - 確認主題設置

4. **功能不工作**
   - 檢查JavaScript錯誤
   - 確認事件監聽器設置
   - 檢查DOM元素是否存在

### 調試技巧

- 使用瀏覽器開發者工具
- 查看控制台錯誤信息
- 檢查網絡請求
- 使用SharePoint工作台測試

## 📞 支持

如果遇到問題，請檢查：
1. SharePoint Framework官方文檔
2. 項目README文件
3. 瀏覽器控制台錯誤信息
4. SharePoint管理中心日誌

## 🎯 下一步

部署成功後，你可以：
1. 自定義學生數據
2. 添加更多功能（排序、分頁等）
3. 連接真實的數據源
4. 優化用戶界面
5. 添加更多統計功能

祝你使用愉快！🎉
