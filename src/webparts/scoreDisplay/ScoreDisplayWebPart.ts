import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './ScoreDisplayWebPart.module.scss';
import * as strings from 'ScoreDisplayWebPartStrings';

export interface IScoreDisplayWebPartProps {
  description: string;
}

interface Student {
  studentId: string;
  name: string;
  chinese: number;
  english: number;
  math: number;
  science: number;
  total: number;
  average: number;
}

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

  private students: Student[] = [];
  private filteredStudents: Student[] = [];
  private pageInfo: PageInfo | null = null;
  private userInfo: UserInfo | null = null;

  public render(): void {
    try {
      this.domElement.innerHTML = `
      <div class="${styles.scoreDisplay || 'scoreDisplay'}">
        <div class="${styles.header || 'header'}">
          <h1>學生成績</h1>
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
                placeholder="搜尋學生姓名或學號..." 
                class="${styles.filterInput || 'filterInput'}"
              >
              <button id="exportBtn" class="${styles.exportBtn || 'exportBtn'}">匯出 CSV</button>
            </div>
            
            <!-- 成績表格 -->
            <div class="${styles.tableContainer || 'tableContainer'}">
              <table id="gradesTable" class="${styles.gradesTable || 'gradesTable'}">
                <thead>
                  <tr>
                    <th>學號</th>
                    <th>姓名</th>
                    <th>中文</th>
                    <th>英文</th>
                    <th>數學</th>
                    <th>科學</th>
                    <th>總分</th>
                    <th>平均分</th>
                  </tr>
                </thead>
                <tbody id="tableBody">
                </tbody>
              </table>
            </div>
            
            <!-- 統計資訊 -->
            <div class="${styles.summaryStats || 'summaryStats'}">
              <div class="${styles.statCard || 'statCard'}">
                <h3>總學生數</h3>
                <p id="totalStudents" class="${styles.statNumber || 'statNumber'}">0</p>
              </div>
              <div class="${styles.statCard || 'statCard'}">
                <h3>平均分數</h3>
                <p id="averageScore" class="${styles.statNumber || 'statNumber'}">0.00</p>
              </div>
              <div class="${styles.statCard || 'statCard'}">
                <h3>最高分數</h3>
                <p id="highestScore" class="${styles.statNumber || 'statNumber'}">0</p>
              </div>
              <div class="${styles.statCard || 'statCard'}">
                <h3>最低分數</h3>
                <p id="lowestScore" class="${styles.statNumber || 'statNumber'}">0</p>
              </div>
            </div>
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
      console.log(`User: ${this.userInfo?.displayName}`);

      // 根據頁面名稱加載不同的模擬數據
      this.students = this._getMockDataByPage(this.pageInfo.pageName);

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
      }, 1000);
    } catch (error) {
      console.error('Error in _loadStudentDataByPage:', error);
      this._showError('載入頁面數據失敗');
    }
  }

  private _getMockDataByPage(pageName: string): Student[] {
    // 根據不同頁面返回不同的模擬數據
    const mockDataMap: { [key: string]: Student[] } = {
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
      { studentId: 'DEFAULT-001', name: '默認學生', chinese: 80, english: 80, math: 80, science: 80, total: 320, average: 80.0 }
    ];
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
      const searchTerm = searchInput?.value.toLowerCase() || '';

      this.filteredStudents = this.students.filter(student => {
        const matchesSearch = student.name.toLowerCase().indexOf(searchTerm) !== -1 || 
                              student.studentId.toLowerCase().indexOf(searchTerm) !== -1;
        return matchesSearch;
      });

      this._renderTable();
      this._updateStats();
    } catch (error) {
      console.error('Error in _filterStudents:', error);
    }
  }

  private _renderTable(): void {
    try {
      const tableBody = this.domElement.querySelector('#tableBody');
      if (!tableBody) return;

      tableBody.innerHTML = this.filteredStudents.map(student => `
        <tr>
          <td>${this._escapeHtml(student.studentId)}</td>
          <td>${this._escapeHtml(student.name)}</td>
          <td>${student.chinese}</td>
          <td>${student.english}</td>
          <td>${student.math}</td>
          <td>${student.science}</td>
          <td>${student.total}</td>
          <td>${student.average.toFixed(2)}</td>
        </tr>
      `).join('');
    } catch (error) {
      console.error('Error in _renderTable:', error);
    }
  }

  private _updateStats(): void {
    try {
      const totalStudents = this.domElement.querySelector('#totalStudents');
      const averageScore = this.domElement.querySelector('#averageScore');
      const highestScore = this.domElement.querySelector('#highestScore');
      const lowestScore = this.domElement.querySelector('#lowestScore');

      if (this.filteredStudents.length === 0) {
        if (totalStudents) totalStudents.textContent = '0';
        if (averageScore) averageScore.textContent = '0.00';
        if (highestScore) highestScore.textContent = '0';
        if (lowestScore) lowestScore.textContent = '0';
        return;
      }

      const avg = this.filteredStudents.reduce((sum, student) => sum + student.average, 0) / this.filteredStudents.length;
      const max = Math.max(...this.filteredStudents.map(student => student.average));
      const min = Math.min(...this.filteredStudents.map(student => student.average));

      if (totalStudents) totalStudents.textContent = this.filteredStudents.length.toString();
      if (averageScore) averageScore.textContent = avg.toFixed(2);
      if (highestScore) highestScore.textContent = max.toString();
      if (lowestScore) lowestScore.textContent = min.toString();
    } catch (error) {
      console.error('Error in _updateStats:', error);
    }
  }

  private _exportToCSV(): void {
    try {
      const pageName = this.pageInfo?.pageName || 'Unknown';
      const headers = ['學號', '姓名', '中文', '英文', '數學', '科學', '總分', '平均分'];
      const csvContent = [
        headers.join(','),
        ...this.filteredStudents.map(student => 
          [student.studentId, student.name, student.chinese, student.english, student.math, student.science, student.total, student.average.toFixed(2)].join(',')
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
