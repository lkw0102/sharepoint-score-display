# 故障排除指南 - 學生成績顯示 Web Part

## 🚨 正式環境錯誤解決方案

### 問題描述
在 `https://groupespauloedu.sharepoint.com/` 正式環境中，Web Part顯示：
```
Something went wrong
If the problem persists, contact the site administrator and give them the information in Technical Details.
Technical Details
ERROR: [object Object]
```

## 🔧 解決方案

### 1. 使用生產環境構建版本

我們已經創建了一個更穩定的生產環境版本，包含：
- 完整的錯誤處理
- 防護性編程
- 樣式類名備用方案
- HTML轉義保護

**文件位置：**
```
student-grades-spfx-new/sharepoint/solution/student-grades-display.sppkg
```

### 2. 重新部署步驟

#### 步驟1: 移除舊版本
1. 進入 SharePoint 管理中心
2. 應用程式目錄 → 管理應用程式
3. 找到舊的 "學生成績顯示" 應用程式
4. 點擊 "移除" 或 "刪除"

#### 步驟2: 上傳新版本
1. 應用程式目錄 → 上傳
2. 選擇新的 `student-grades-display.sppkg` 文件
3. 點擊 "部署"

#### 步驟3: 重新添加Web Part
1. 進入SharePoint頁面
2. 編輯頁面
3. 添加新的 "學生成績顯示" Web Part

### 3. 常見錯誤原因

#### 3.1 CSS樣式問題
**症狀：** Web Part顯示但樣式異常
**解決方案：** 新版本已添加樣式類名備用方案

#### 3.2 JavaScript錯誤
**症狀：** 功能不工作或顯示錯誤
**解決方案：** 新版本包含完整的錯誤處理

#### 3.3 DOM操作失敗
**症狀：** 搜索、篩選等功能不響應
**解決方案：** 新版本添加了防護性DOM檢查

### 4. 調試技巧

#### 4.1 瀏覽器開發者工具
1. 按 F12 打開開發者工具
2. 查看 Console 標籤的錯誤信息
3. 查看 Network 標籤的網絡請求

#### 4.2 SharePoint工作台測試
```
https://groupespauloedu.sharepoint.com/_layouts/workbench.aspx
```

#### 4.3 檢查Web Part屬性
1. 編輯Web Part
2. 檢查屬性面板設置
3. 確保沒有特殊字符或無效配置

### 5. 環境特定問題

#### 5.1 權限問題
- 確保用戶有權限訪問Web Part
- 檢查SharePoint List權限（如果使用真實數據）

#### 5.2 瀏覽器兼容性
- 確保使用現代瀏覽器（Chrome, Edge, Firefox）
- 清除瀏覽器緩存

#### 5.3 網絡問題
- 檢查網絡連接
- 確認SharePoint服務正常

### 6. 應急解決方案

如果Web Part仍然無法正常工作，可以：

#### 6.1 使用簡單版本
創建一個最小化的Web Part版本，只顯示基本表格：

```typescript
public render(): void {
  this.domElement.innerHTML = `
    <div style="padding: 20px; font-family: 'Segoe UI', sans-serif;">
      <h2>學生成績系統</h2>
      <table style="width: 100%; border-collapse: collapse;">
        <thead>
          <tr style="background-color: #f3f2f1;">
            <th style="padding: 8px; border: 1px solid #ddd;">學號</th>
            <th style="padding: 8px; border: 1px solid #ddd;">姓名</th>
            <th style="padding: 8px; border: 1px solid #ddd;">中文</th>
            <th style="padding: 8px; border: 1px solid #ddd;">英文</th>
            <th style="padding: 8px; border: 1px solid #ddd;">數學</th>
            <th style="padding: 8px; border: 1px solid #ddd;">科學</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;">S001</td>
            <td style="padding: 8px; border: 1px solid #ddd;">張小明</td>
            <td style="padding: 8px; border: 1px solid #ddd;">85</td>
            <td style="padding: 8px; border: 1px solid #ddd;">92</td>
            <td style="padding: 8px; border: 1px solid #ddd;">88</td>
            <td style="padding: 8px; border: 1px solid #ddd;">90</td>
          </tr>
        </tbody>
      </table>
    </div>
  `;
}
```

#### 6.2 使用HTML嵌入
如果SPFx Web Part持續有問題，可以考慮使用HTML嵌入方式：

1. 創建一個簡單的HTML文件
2. 上傳到SharePoint文檔庫
3. 使用"嵌入" Web Part來顯示

### 7. 聯繫支持

如果問題持續存在，請提供以下信息：

1. **錯誤截圖** - 完整的錯誤信息
2. **瀏覽器信息** - 瀏覽器類型和版本
3. **控制台日誌** - 開發者工具中的錯誤信息
4. **環境信息** - SharePoint版本和配置

### 8. 預防措施

#### 8.1 測試環境
- 在測試環境中先驗證Web Part
- 使用不同的瀏覽器和設備測試

#### 8.2 版本控制
- 保留舊版本的備份
- 記錄每次更改的內容

#### 8.3 監控
- 定期檢查Web Part功能
- 監控用戶反饋

## 📞 緊急聯繫

如果遇到緊急問題：
1. 檢查本指南的解決方案
2. 查看瀏覽器控制台錯誤
3. 嘗試重新部署Web Part
4. 聯繫技術支持團隊

---

**注意：** 新版本已經過優化，應該能夠解決大部分正式環境中的問題。如果問題持續，請使用應急解決方案或聯繫支持團隊。
