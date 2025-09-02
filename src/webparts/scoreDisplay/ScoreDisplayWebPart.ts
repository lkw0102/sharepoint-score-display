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
  excelFilePath: string;
  headerTitle?: string; // Optional custom header title
}

type StudentRow = Record<string, string | number | undefined>;

interface PageInfo {
  pageName: string;
  siteUrl: string;
  listTitle?: string;
  excelFileName?: string;
}

interface UserInfo {
  id: number;
  displayName: string;
  email: string;
  loginName: string;
  isSiteAdmin: boolean;
  roles: string[];
  permissions: string[];
  groups: string[];
  rawUserData?: any; // 保存原始 API 數據
  rawGroupsData?: any; // 保存原始群組數據
  rawPermissionsData?: any; // 保存原始權限數據
}

export default class ScoreDisplayWebPart extends BaseClientSideWebPart<IScoreDisplayWebPartProps> {

  private students: StudentRow[] = [];
  private filteredStudents: StudentRow[] = [];
  private columns: string[] = [];
  private pageInfo: PageInfo | null = null;
  private userInfo: UserInfo | null = null;
  private emailPrefix: string = '';

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
                                 <!-- 篩選器 - 暫時隱藏 -->
                     <!-- <div class="${styles.filters || 'filters'}">
                       <input 
                         id="searchInput"
                         type="text" 
                         placeholder="搜尋...（跨所有欄位）" 
                         class="${styles.filterInput || 'filterInput'}"
                       >
                     </div> -->
            
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
      this.emailPrefix = this._extractEmailPrefix(this.userInfo.email);
      // this._logUserInfo();

      // 根據頁面信息加載對應的學生數據
      await this._loadStudentData();
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
      // 並行獲取用戶基本信息、群組和權限
      const [userResponse, groupsResponse, rolesResponse] = await Promise.all([
        this.context.spHttpClient.get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser`,
          SPHttpClient.configurations.v1
        ),
        this.context.spHttpClient.get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser/groups`,
          SPHttpClient.configurations.v1
        ),
        this.context.spHttpClient.get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/getusereffectivepermissions(@v)?@v='${encodeURIComponent(this.context.pageContext.user.loginName)}'`,
          SPHttpClient.configurations.v1
        )
      ]);

      const userData = userResponse.ok ? await userResponse.json() : null;
      const groupsData = groupsResponse.ok ? await groupsResponse.json() : { value: [] };
      const rolesData = rolesResponse.ok ? await rolesResponse.json() : null;

      // 提取群組信息
      const groups = groupsData.value?.map((group: any) => group.Title) || [];
      
      // 解析權限
      const permissions = this._parseUserPermissions(rolesData);
      
      // 判斷用戶角色
      const roles = this._determineUserRoles(groups, userData?.IsSiteAdmin, permissions);

      return {
        id: userData?.Id || 0,
        displayName: userData?.Title || userData?.DisplayName || this.context.pageContext.user.displayName,
        email: userData?.Email || this.context.pageContext.user.email,
        loginName: userData?.LoginName || this.context.pageContext.user.loginName,
        isSiteAdmin: userData?.IsSiteAdmin || false,
        roles: roles,
        permissions: permissions,
        groups: groups,
        rawUserData: userData, // 保存完整的原始用戶數據
        rawGroupsData: groupsData, // 保存完整的原始群組數據
        rawPermissionsData: rolesData // 保存完整的原始權限數據
      };
    } catch (error) {
      console.error('Error getting user info:', error);
      // 返回基本信息
      return {
        id: 0,
        displayName: this.context.pageContext.user.displayName,
        email: this.context.pageContext.user.email,
        loginName: this.context.pageContext.user.loginName,
        isSiteAdmin: false,
        roles: ['Guest'],
        permissions: [],
        groups: [],
        rawUserData: null,
        rawGroupsData: null,
        rawPermissionsData: null
      };
    }
  }

  private _parseUserPermissions(rolesData: any): string[] {
    const permissions: string[] = [];
    if (!rolesData || !rolesData.GetUserEffectivePermissions) return permissions;

    const permissionValue = rolesData.GetUserEffectivePermissions;
    
    // SharePoint 權限位元對應 - 使用英文原始名稱
    const permissionMap: { [key: number]: string } = {
      1: 'ViewListItems',
      2: 'AddListItems', 
      4: 'EditListItems',
      8: 'DeleteListItems',
      16: 'ApproveItems',
      32: 'OpenItems',
      64: 'ViewVersions',
      128: 'DeleteVersions',
      256: 'CreateAlerts',
      512: 'ViewFormPages',
      1024: 'AnonymousSearchAccessList',
      2048: 'UseRemoteAPIs',
      4096: 'UseClientIntegration',
      8192: 'UseSelfServiceSiteCreation',
      16384: 'UsePersonalFeatures',
      32768: 'AddDelPrivateWebParts',
      65536: 'UpdatePersonalWebParts',
      131072: 'ManageLists',
      262144: 'ManageWeb',
      524288: 'CreateSubwebs',
      1048576: 'ManagePermissions',
      2097152: 'BrowseDirectories',
      4194304: 'BrowseUserInfo',
      8388608: 'AddAndCustomizePages',
      16777216: 'ApplyThemeAndBorder',
      33554432: 'ApplyStyleSheets',
      67108864: 'CreateGroups',
      134217728: 'BrowseWeb',
      268435456: 'UseSSO',
      536870912: 'EditMyUserInfo',
      1073741824: 'EnumeratePermissions'
    };

    // 解析權限位元
    for (const bit in permissionMap) {
      if (Object.prototype.hasOwnProperty.call(permissionMap, bit)) {
        if (permissionValue & parseInt(bit)) {
          permissions.push(permissionMap[parseInt(bit)]);
        }
      }
    }

    return permissions;
  }

  private _determineUserRoles(groups: string[], isSiteAdmin: boolean, permissions: string[]): string[] {
    const roles: string[] = [];

    // 直接返回 SharePoint 群組名稱，不進行中文轉換
    if (isSiteAdmin) {
      roles.push('Site Administrator');
    }

    // 直接使用群組名稱
    for (const group of groups) {
      roles.push(group);
    }

    return roles;
  }

  // @ts-expect-error - Method is used for debugging purposes
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  private _logUserInfo(): void {
    try {
      if (!this.userInfo) {
        console.log('UserInfo is not available');
        return;
      }
      console.group('SPFx UserInfo');
      console.log('userInfo:', this.userInfo);
      console.log('=== 原始 API 數據 ===');
      console.log('rawUserData:', this.userInfo.rawUserData);
      console.log('rawGroupsData:', this.userInfo.rawGroupsData);
      console.log('rawPermissionsData:', this.userInfo.rawPermissionsData);
      console.log('=== 提取的 Email 信息 ===');
      console.log('email:', this.userInfo.email);
      console.log('emailPrefix (全域變數):', this.emailPrefix);
      console.groupEnd();
    } catch (error) {
      console.error('Error logging user info:', error);
    }
  }

  private async _loadStudentData(): Promise<void> {
    try {
      if (!this.pageInfo) {
        throw new Error('Page info not available');
      }

      // 從用戶設定的 Excel 檔案路徑載入資料
      if (!this.properties.excelFilePath) {
        this._showError('請在 Web Part 屬性中設定 Excel 檔案路徑');
        return;
      }
      await this._loadStudentDataFromExcel(this.properties.excelFilePath);
      this._updateHeaderTitle(this.properties.excelFilePath);

      // 依據資料動態生成欄位
      this.columns = this._deriveColumns(this.students);

      // 模擬載入延遲
      setTimeout(() => {
        try {
          // 立即執行篩選，用戶只能看到自己的資料
          this._filterStudents();
          this._hideLoading();
          this._renderTable();
          this._updateStats();
        } catch (error) {
          console.error('Error in data loading timeout:', error);
          this._showError('數據載入失敗');
        }
      }, 500);
    } catch (error) {
      console.error('Error in _loadStudentData:', error);
      this._showError('載入頁面數據失敗');
    }
  }

  private async _loadStudentDataFromExcel(filePath: string): Promise<void> {
    try {
      let apiUrl: string;
      const origin = window.location.origin;

      const getWebBaseFromPath = (pathOrUrl: string): string => {
        try {
          // Normalize to absolute URL to parse
          const asUrl = new URL(pathOrUrl, origin);
          const pathname = asUrl.pathname; // e.g., /sites/Classrooms/Shared Documents/...
          // Handle common managed paths: /sites/{site}, /teams/{team}
          const segments = pathname.split('/').filter(seg => seg);
          const managedRoots = ['sites', 'teams'];
          if (segments.length >= 2 && managedRoots.indexOf(segments[0]) !== -1) {
            // Return origin + /sites/{site}
            return `${asUrl.origin}/${segments[0]}/${segments[1]}`;
          }
          // Fallback to current web
          return this.context.pageContext.web.absoluteUrl;
        } catch {
          return this.context.pageContext.web.absoluteUrl;
        }
      };
      
      // 判斷是否為 SharePoint 共享連結格式
      if (filePath.indexOf('sharepoint.com/:x:/') !== -1) {
        // 檢查是否為 /s/ 格式（需要特殊處理）
        if (filePath.indexOf('sharepoint.com/:x:/s/') !== -1) {
          // 使用共享連結 API
          // Use tenant origin for v2 shares API to avoid scoping to current web
          apiUrl = `${origin}/_api/v2.0/shares/u!${btoa(filePath).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-')}/driveItem/content`;
        } else {
          // /r/ 格式，提取伺服器相對路徑
          const serverRelativePath = this._extractServerRelativePathFromShareLink(filePath);
          if (!serverRelativePath) {
            throw new Error('無法從共享連結提取檔案路徑');
          }
          const encodedPath = encodeURI(serverRelativePath);
          const webBase = getWebBaseFromPath(serverRelativePath);
          apiUrl = `${webBase}/_api/web/GetFileByServerRelativeUrl('${encodedPath}')/$value`;
        }
      } else if (filePath.indexOf('_layouts/15/doc.aspx') !== -1) {
        // SharePoint 文檔編輯連結格式
        const docInfo = this._extractDocInfoFromEditLink(filePath);
        if (!docInfo) {
          throw new Error('無法從文檔編輯連結提取檔案信息');
        }
        // Attempt to call against the file's web if URL reveals it; fall back to current web
        const webBase = getWebBaseFromPath(filePath);
        apiUrl = `${webBase}/_api/web/GetFileById('${docInfo.fileId}')/$value`;
      } else {
        // 假設為伺服器相對路徑
        const encodedPath = encodeURI(filePath);
        const webBase = getWebBaseFromPath(filePath);
        apiUrl = `${webBase}/_api/web/GetFileByServerRelativeUrl('${encodedPath}')/$value`;
      }
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (!response.ok) {
        const text = await response.text().catch(() => '');
        throw new Error(`Failed to download Excel: ${response.status} ${response.statusText} ${text}`);
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
      // 暫時隱藏搜索功能
      // const searchInput = this.domElement.querySelector('#searchInput') as HTMLInputElement;
      // if (searchInput) {
      //   searchInput.addEventListener('input', () => this._filterStudents());
      // }
    } catch (error) {
      console.error('Error setting up event listeners:', error);
    }
  }

  private _filterStudents(): void {
    try {
      const searchInput = this.domElement.querySelector('#searchInput') as HTMLInputElement;
      const searchTerm = (searchInput?.value || '').toLowerCase();

      // 首先根據 emailPrefix 篩選學生帳號
      let filteredByAccount = this.students;
      if (this.emailPrefix) {
        const studentAccountField = this._findStudentAccountField();
        
        filteredByAccount = this.students.filter(row => {
          if (studentAccountField && row[studentAccountField]) {
            const studentAccount = String(row[studentAccountField]).toLowerCase();
            return studentAccount === this.emailPrefix.toLowerCase();
          }
          return false;
        });
      }

      // 然後根據搜尋詞進一步篩選
      if (!searchTerm) {
        this.filteredStudents = filteredByAccount;
      } else {
        this.filteredStudents = filteredByAccount.filter(row =>
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

  private _extractDocInfoFromEditLink(editLink: string): { fileId: string; fileName?: string } | null {
    try {
      // 從 SharePoint 文檔編輯連結提取檔案信息
      // 格式：https://domain.sharepoint.com/sites/sitename/_layouts/15/doc.aspx?sourcedoc={GUID}&action=edit
      
      // 提取 sourcedoc 參數（檔案 GUID）
      const sourcedocMatch = editLink.match(/sourcedoc=([^&]+)/);
      if (!sourcedocMatch || !sourcedocMatch[1]) {
        return null;
      }
      
      const fileId = sourcedocMatch[1].replace(/[{}]/g, ''); // 移除大括號
      
      // 嘗試提取檔案名稱（如果有的話）
      const fileNameMatch = editLink.match(/&file=([^&]+)/);
      const fileName = fileNameMatch ? decodeURIComponent(fileNameMatch[1]) : undefined;
      
      return { fileId, fileName };
    } catch (error) {
      console.error('Error extracting doc info from edit link:', error);
      return null;
    }
  }

  private _extractServerRelativePathFromShareLink(shareLink: string): string | null {
    try {
      // 從 SharePoint 共享連結提取伺服器相對路徑
      // 格式1：https://domain.sharepoint.com/:x:/r/sites/sitename/Shared%20Documents/path/file.xlsx
      // 格式2：https://domain.sharepoint.com/:x:/s/sitename/EbGQlGpsdeNPuN0NXEl1AHQBsGdfG5WFZPEAIXbR4AySpQ?e=gXZ1NW
      
      // 嘗試格式1 (/r/ 格式)
      let match = shareLink.match(/sharepoint\.com\/:x:\/r(\/sites\/[^?]+)/);
      if (match && match[1]) {
        return decodeURIComponent(match[1]);
      }
      
      // 嘗試格式2 (/s/ 格式) - 這種格式需要不同的處理方式
      // eslint-disable-next-line no-useless-escape
      match = shareLink.match(/sharepoint\.com\/:x:\/s\/([^\/]+)/);
      if (match && match[1]) {
        // 對於 /s/ 格式，我們無法直接從 URL 得到檔案路徑
        // 需要使用不同的 API 方法
        return null; // 標記為需要特殊處理
      }
      
      return null;
    } catch (error) {
      console.error('Error extracting server relative path:', error);
      return null;
    }
  }

  private _updateHeaderTitle(excelPath: string | null): void {
    try {
      const headerTitle = this.domElement.querySelector('#headerTitle') as HTMLElement;
      if (!headerTitle) return;

      // 優先使用用戶設定的自定義標題
      if (this.properties.headerTitle && this.properties.headerTitle.trim()) {
        headerTitle.textContent = this.properties.headerTitle.trim();
        return;
      }

      // 如果沒有自定義標題，則從檔案路徑提取
      if (excelPath && excelPath.trim()) {
        // 從路徑提取檔名並移除 .xlsx
        let fileName = '';
        if (excelPath.indexOf('sharepoint.com/:x:/') !== -1) {
          if (excelPath.indexOf('sharepoint.com/:x:/s/') !== -1) {
            // /s/ 格式無法從 URL 得到檔名，使用通用名稱
            fileName = 'Excel 檔案';
          } else {
            // /r/ 格式，從共享連結提取檔名
            const serverPath = this._extractServerRelativePathFromShareLink(excelPath);
            fileName = serverPath?.split('/').pop() || 'Excel 檔案';
          }
        } else if (excelPath.indexOf('_layouts/15/doc.aspx') !== -1) {
          // 文檔編輯連結格式
          const docInfo = this._extractDocInfoFromEditLink(excelPath);
          fileName = docInfo?.fileName || 'Excel 檔案';
        } else {
          fileName = excelPath.split('/').pop() || '';
        }
        const displayName = fileName.replace(/\.xlsx$/i, '');
        headerTitle.textContent = displayName;
      } else {
        // 如果沒有 Excel 路徑，使用頁面名稱或預設標題
        headerTitle.textContent = this.pageInfo?.pageName || '學生成績';
      }
    } catch (error) {
      console.error('Error updating header title:', error);
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

  private _findStudentAccountField(): string | null {
    // 常見的學生帳號欄位名稱
    const possibleFieldNames = [
      '學生帳號'
    ];

    for (const fieldName of possibleFieldNames) {
      if (this.columns.indexOf(fieldName) !== -1) {
        return fieldName;
      }
    }

    // 如果找不到，返回第一個欄位作為備選
    return this.columns.length > 0 ? this.columns[0] : null;
  }

  private _extractEmailPrefix(email: string): string {
    if (!email || typeof email !== 'string') {
      return '';
    }
    const atIndex = email.indexOf('@');
    if (atIndex === -1) {
      return email; // 如果沒有 @ 符號，返回整個 email
    }
    return email.substring(0, atIndex);
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
                PropertyPaneTextField('headerTitle', {
                  label: '標題',
                  description: '可選：設定自定義的標題，留空則自動使用檔案名稱',
                  placeholder: '例如：F1A代數'
                }),
                PropertyPaneTextField('excelFilePath', {
                  label: 'Excel 檔案路徑',
                  description: '輸入Excel的Link，例如：https://groupespauloedu.sharepoint.com/:x:/r/sites/Classrooms/Shared%20Documents/F1/F1_head.xlsx',
                  placeholder: 'https://groupespauloedu.sharepoint.com/..../f1.xlsx'
                }),
     
              ]
            }
          ]
        }
      ]
    };
  }
}
