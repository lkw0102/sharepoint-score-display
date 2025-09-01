import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as XLSX from 'xlsx';

import styles from './ScoreDisplayWebPart.module.scss';
import * as strings from 'ScoreDisplayWebPartStrings';

export interface IScoreDisplayWebPartProps {
  description: string;
}

type StudentRow = Record<string, string | number | undefined>;

interface PageInfo {
  pageName: string;
  siteUrl: string;
  listTitle?: string;
  excelFileName?: string;
}

interface UserInfo {
  displayName: string;
  email: string;
  loginName: string;
  isSiteAdmin: boolean;
}

export default class ScoreDisplayWebPart extends BaseClientSideWebPart<IScoreDisplayWebPartProps> {

  private students: StudentRow[] = [];
  private filteredStudents: StudentRow[] = [];
  private columns: string[] = [];
  private pageInfo: PageInfo | null = null;
  private userInfo: UserInfo | null = null;

  public render(): void {
    try {
      this.domElement.innerHTML = `
      <div class="${styles.scoreDisplay || 'scoreDisplay'}">
        <div class="${styles.header || 'header'}">
          <h1 id="headerTitle">載入中...</h1>
        </div>
        
        <div class="${styles.content || 'content'}">
          <div id="loading" class="${styles.loading || 'loading'}">
            <div class="${styles.spinner || 'spinner'}"></div>
            <p>正在載入學生成績資料...</p>
          </div>
          
          <div id="error" class="${styles.error || 'error'}" style="display: none;">
            <p id="errorMessage">載入失敗，請稍後再試</p>
          </div>
          
          <div id="mainContent" style="display: none;">
            <!-- 篩選器 -->
            <div class="${styles.filters || 'filters'}">
              <input 
                id="searchInput"
                type="text" 
                placeholder="搜尋...（跨所有欄位）" 
                class="${styles.filterInput || 'filterInput'}"
              >
              <button id="exportBtn" class="${styles.exportBtn || 'exportBtn'}">匯出 CSV</button>
            </div>
            
            <!-- 成績表格 -->
            <div class="${styles.tableContainer || 'tableContainer'}">
              <table id="gradesTable" class="${styles.gradesTable || 'gradesTable'}">
                <thead>
                  <tr id="tableHeader"></tr>
                </thead>
                <tbody id="tableBody"></tbody>
              </table>
            </div>
            
            <!-- 統計資訊已移除（依需求不顯示） -->
          </div>
        </div>
      </div>`;

      this._initializeData().catch(error => {
        console.error('Error initializing data:', error);
        this._showError('初始化失敗');
      });
      this._setupEventListeners();
    } catch (error) {
      console.error('Error in render method:', error);
      this.domElement.innerHTML = `
        <div style="padding: 20px; text-align: center; color: #d13438;">
          <h2>載入失敗</h2>
          <p>學生成績系統載入時發生錯誤，請稍後再試。</p>
          <p style="font-size: 12px; color: #666;">錯誤詳情: ${error.message}</p>
        </div>
      `;
    }
  }

  private async _initializeData(): Promise<void> {
    try {
      // 並行獲取頁面信息和用戶信息
      const [pageInfo, userInfo] = await Promise.all([
        this._getPageInfo(),
        this._getUserInfo()
      ]);

      this.pageInfo = pageInfo;
      this.userInfo = userInfo;
      this._logUserInfo();
      if (this.pageInfo) {
        console.log('Current page:', this.pageInfo.pageName, '| Site URL:', this.pageInfo.siteUrl);
      }

      // 根據頁面信息加載對應的學生數據
      await this._loadStudentDataByPage();
    } catch (error) {
      console.error('Error in _initializeData:', error);
      this._showError('初始化數據失敗');
    }
  }

  private async _getPageInfo(): Promise<PageInfo> {
    try {
      const currentUrl = window.location.href;
      const urlParts = currentUrl.split('/');
      const pageName = urlParts[urlParts.length - 1].replace('.aspx', '');
      
      return {
        pageName: pageName,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        listTitle: `成績資料_${pageName}`,
        excelFileName: `學生成績_${pageName}.xlsx`
      };
    } catch (error) {
      console.error('Error getting page info:', error);
      return {
        pageName: 'Unknown',
        siteUrl: this.context.pageContext.web.absoluteUrl
      };
    }
  }

  private async _getUserInfo(): Promise<UserInfo> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const userData = await response.json();
        return {
          displayName: userData.Title || userData.DisplayName,
          email: userData.Email,
          loginName: userData.LoginName,
          isSiteAdmin: userData.IsSiteAdmin || false
        };
      } else {
        throw new Error(`Failed to get user info: ${response.statusText}`);
      }
    } catch (error) {
      console.error('Error getting user info:', error);
      // 返回基本信息
      return {
        displayName: this.context.pageContext.user.displayName,
        email: this.context.pageContext.user.email,
        loginName: this.context.pageContext.user.loginName,
        isSiteAdmin: false
      };
    }
  }

  private _logUserInfo(): void {
    try {
      if (!this.userInfo) {
        console.log('UserInfo is not available');
        return;
      }
      console.group('SPFx UserInfo');
      console.log('userInfo:', this.userInfo);
      console.log('displayName:', this.userInfo.displayName);
      console.log('email:', this.userInfo.email);
      console.log('loginName:', this.userInfo.loginName);
      console.log('isSiteAdmin:', this.userInfo.isSiteAdmin);
      console.groupEnd();
    } catch (error) {
      console.error('Error logging user info:', error);
    }
  }

  private async _loadStudentDataByPage(): Promise<void> {
    try {
      if (!this.pageInfo) {
        throw new Error('Page info not available');
      }

      console.log(`Loading data for page: ${this.pageInfo.pageName}`);

      // 當前頁面 TEST1: 從實際 Excel 檔載入
      if (this.pageInfo.pageName.toUpperCase() === 'TEST1') {
        const serverRelativePath = '/sites/Classrooms/Shared Documents/F1/F1A/F1A_代數.xlsx';
        await this._loadStudentDataFromExcel(serverRelativePath);
        this._updateHeaderTitle(serverRelativePath);
      } else {
        // 其他頁面使用模擬資料
        this.students = this._getMockDataByPage(this.pageInfo.pageName);
        this._updateHeaderTitle(null);
      }

      // 依據資料動態生成欄位
      this.columns = this._deriveColumns(this.students);

      // 模擬載入延遲
      setTimeout(() => {
        try {
          this.filteredStudents = [...this.students];
          this._hideLoading();
          this._renderTable();
          this._updateStats();
        } catch (error) {
          console.error('Error in data loading timeout:', error);
          this._showError('數據載入失敗');
        }
      }, 500);
    } catch (error) {
      console.error('Error in _loadStudentDataByPage:', error);
      this._showError('載入頁面數據失敗');
    }
  }

  private async _loadStudentDataFromExcel(serverRelativePath: string): Promise<void> {
    try {
      // 下載 Excel 檔為 ArrayBuffer
      const encodedPath = encodeURI(serverRelativePath);
      const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${encodedPath}')/$value`;
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (!response.ok) {
        throw new Error(`Failed to download Excel: ${response.status} ${response.statusText}`);
      }
      const arrayBuffer = await response.arrayBuffer();

      // 使用 SheetJS 解析
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      if (!firstSheetName) {
        throw new Error('Workbook has no sheets');
      }
      const sheet = workbook.Sheets[firstSheetName];

      // 以 header:1 轉為二維陣列，第一列為表頭
      const rows: unknown[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as unknown[][];
      if (!rows || rows.length === 0) {
        this.students = [];
        return;
      }

      // 取得表頭
      const headerRow = (rows[0] as (string | number | null | undefined)[]).map((h, idx) => {
        const headerStr = h === null || h === undefined ? '' : String(h).trim();
        return headerStr.length === 0 ? `Column${idx + 1}` : headerStr;
      });

      // 其餘資料列轉為物件陣列
      const dataObjects: StudentRow[] = [];
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r] as (string | number | null | undefined)[];
        if (!row || row.length === 0) { continue; }
        const obj: StudentRow = {};
        for (let c = 0; c < headerRow.length; c++) {
          const key = headerRow[c];
          const value = row[c];
          if (value === null || value === undefined) {
            obj[key] = undefined;
          } else if (typeof value === 'number') {
            obj[key] = value;
          } else {
            // 嘗試將數字字串轉為數字
            const num = Number(String(value).trim());
            obj[key] = isNaN(num) ? String(value) : num;
          }
        }
        dataObjects.push(obj);
      }

      this.students = dataObjects;
      this.columns = headerRow;
    } catch (error) {
      console.error('Error in _loadStudentDataFromExcel:', error);
      this._showError('讀取 Excel 失敗');
      // 回退為空資料，避免中斷
      this.students = [];
      this.columns = [];
    }
  }

  private _getMockDataByPage(pageName: string): StudentRow[] {
    // 根據不同頁面返回不同的模擬數據
    const mockDataMap: { [key: string]: StudentRow[] } = {
      'TEST1': [
        { studentId: 'T1-001', name: '張小明', chinese: 85, english: 92, math: 88, science: 90, total: 355, average: 88.75 },
        { studentId: 'T1-002', name: '李小華', chinese: 78, english: 85, math: 92, science: 87, total: 342, average: 85.5 },
        { studentId: 'T1-003', name: '王小美', chinese: 92, english: 88, math: 85, science: 94, total: 359, average: 89.75 }
      ],
      'TEST2': [
        { studentId: 'T2-001', name: '陳小強', chinese: 80, english: 90, math: 95, science: 82, total: 347, average: 86.75 },
        { studentId: 'T2-002', name: '林小芳', chinese: 88, english: 85, math: 90, science: 88, total: 351, average: 87.75 },
        { studentId: 'T2-003', name: '黃小偉', chinese: 75, english: 92, math: 87, science: 90, total: 344, average: 86.0 }
      ],
      'ClassA': [
        { studentId: 'CA-001', name: '劉小玲', chinese: 90, english: 88, math: 93, science: 85, total: 356, average: 89.0 },
        { studentId: 'CA-002', name: '吳小傑', chinese: 82, english: 85, math: 88, science: 92, total: 347, average: 86.75 },
        { studentId: 'CA-003', name: '趙小雅', chinese: 95, english: 90, math: 89, science: 91, total: 365, average: 91.25 }
      ],
      'ClassB': [
        { studentId: 'CB-001', name: '孫小龍', chinese: 87, english: 93, math: 91, science: 89, total: 360, average: 90.0 },
        { studentId: 'CB-002', name: '周小紅', chinese: 83, english: 87, math: 86, science: 93, total: 349, average: 87.25 },
        { studentId: 'CB-003', name: '鄭小剛', chinese: 91, english: 89, math: 94, science: 86, total: 360, average: 90.0 }
      ]
    };

    // 如果沒有特定頁面的數據，返回默認數據
    return mockDataMap[pageName] || [
      { studentId: 'DEFAULT-001', name: '默認學生', chinese1: 80, english: 80, math: 80, science: 80, total: 320, average: 80.0 }
    ];
  }

  private _deriveColumns(rows: StudentRow[]): string[] {
    if (!rows || rows.length === 0) return [];
    const keys = Object.keys(rows[0]);
    // 嘗試以更自然的順序排列常見欄位
    const preferredOrder = ['studentId', 'name'];
    const ordered: string[] = [];
    preferredOrder.forEach(k => { if (keys.indexOf(k) !== -1) ordered.push(k); });
    keys.forEach(k => { if (ordered.indexOf(k) === -1) ordered.push(k); });
    return ordered;
  }

  private _setupEventListeners(): void {
    try {
      const searchInput = this.domElement.querySelector('#searchInput') as HTMLInputElement;
      const exportBtn = this.domElement.querySelector('#exportBtn') as HTMLButtonElement;

      if (searchInput) {
        searchInput.addEventListener('input', () => this._filterStudents());
      }
      if (exportBtn) {
        exportBtn.addEventListener('click', () => this._exportToCSV());
      }
    } catch (error) {
      console.error('Error setting up event listeners:', error);
    }
  }

  private _filterStudents(): void {
    try {
      const searchInput = this.domElement.querySelector('#searchInput') as HTMLInputElement;
      const searchTerm = (searchInput?.value || '').toLowerCase();

      if (!searchTerm) {
        this.filteredStudents = [...this.students];
      } else {
        this.filteredStudents = this.students.filter(row =>
          this.columns.some(col => {
            const value = row[col];
            if (value === null || value === undefined) return false;
            return String(value).toLowerCase().indexOf(searchTerm) !== -1;
          })
        );
      }

      this._renderTable();
      this._updateStats();
    } catch (error) {
      console.error('Error in _filterStudents:', error);
    }
  }

  private _renderTable(): void {
    try {
      const headerRow = this.domElement.querySelector('#tableHeader') as HTMLElement;
      const tableBody = this.domElement.querySelector('#tableBody') as HTMLElement;
      if (!headerRow || !tableBody) return;

      // Render header
      headerRow.innerHTML = this.columns.map(col => `<th>${this._escapeHtml(col)}</th>`).join('');

      // Render rows
      tableBody.innerHTML = this.filteredStudents.map(row => {
        const cells = this.columns.map(col => {
          const value = row[col];
          if (typeof value === 'number') {
            return `<td>${value}</td>`;
          }
          return `<td>${this._escapeHtml(value !== undefined && value !== null ? String(value) : '')}</td>`;
        }).join('');
        return `<tr>${cells}</tr>`;
      }).join('');
    } catch (error) {
      console.error('Error in _renderTable:', error);
    }
  }

  private _updateStats(): void {
    // 按需求不顯示統計資訊，保留空方法以避免呼叫錯誤
    return;
  }

  private _updateHeaderTitle(excelPath: string | null): void {
    try {
      const headerTitle = this.domElement.querySelector('#headerTitle') as HTMLElement;
      if (!headerTitle) return;

      if (excelPath) {
        // 從路徑提取檔名並移除 .xlsx
        const fileName = excelPath.split('/').pop() || '';
        const displayName = fileName.replace(/\.xlsx$/i, '');
        headerTitle.textContent = displayName;
      } else {
        // 其他頁面使用頁面名稱
        headerTitle.textContent = this.pageInfo?.pageName || '學生成績';
      }
    } catch (error) {
      console.error('Error updating header title:', error);
    }
  }

  // 已移除統計功能，不需要判定數值欄位的輔助函式
  private _exportToCSV(): void {
    try {
      const pageName = this.pageInfo?.pageName || 'Unknown';
      const headers = this.columns;
      const csvContent = [
        headers.join(','),
        ...this.filteredStudents.map(row => 
          this.columns.map(col => {
            const v = row[col];
            const s = v === null || v === undefined ? '' : String(v);
            return s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1
              ? '"' + s.replace(/"/g, '""') + '"'
              : s;
          }).join(',')
        )
      ].join('\n');

      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', `student_grades_${pageName}.csv`);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      console.error('Error in _exportToCSV:', error);
      alert('匯出失敗，請稍後再試');
    }
  }

  private _hideLoading(): void {
    try {
      const loading = this.domElement.querySelector('#loading') as HTMLElement;
      const mainContent = this.domElement.querySelector('#mainContent') as HTMLElement;
      
      if (loading) loading.style.display = 'none';
      if (mainContent) mainContent.style.display = 'block';
    } catch (error) {
      console.error('Error in _hideLoading:', error);
    }
  }

  private _showError(message: string): void {
    try {
      const errorDiv = this.domElement.querySelector('#error') as HTMLElement;
      const errorMessage = this.domElement.querySelector('#errorMessage') as HTMLElement;
      const loading = this.domElement.querySelector('#loading') as HTMLElement;
      
      if (loading) loading.style.display = 'none';
      if (errorDiv) errorDiv.style.display = 'block';
      if (errorMessage) errorMessage.textContent = message;
    } catch (error) {
      console.error('Error in _showError:', error);
    }
  }

  private _escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
